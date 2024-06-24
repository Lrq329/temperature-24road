using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using test11.Tools;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using static test11.Tools.CRC16MODBUS;
using Newtonsoft.Json.Linq;
using System.Management;
using System.Net.Sockets;


namespace test11
{
    public partial class Form1 : Form
    {
        // 进制转换类方法实例
        Base_conversion base_Conversion = new Base_conversion();
        // CRC16类方法实例
        CRC16MODBUS.CRC16 crc16 = new CRC16();
        //判断合法方法
        illegal illegal = new illegal();

        // TextBox(1) 温度数组
        private TextBox[] temperatureTextBoxes;
        // TextBox(2) 温度数组
        private TextBox[] temperatureTextBoxes2;
        //TextBox 别名数组
        private TextBox[] tempeartureOtherNameBoxes;
        private TextBox[] tempeartureOtherNameBoxes2;
        //checkBox 图表
        private CheckBox[] temperatureCheckBoxes;
        private CheckBox[] temperatureCheckBoxes2;
        // 记录温度和时间戳
        private Queue<(DateTime timeStamp, List<float> temperatures)> temperatureData = new Queue<(DateTime timeStamp, List<float> temperatures)>();
        private Queue<(DateTime timeStamp, List<float> temperatures)> temperatureData2 = new Queue<(DateTime timeStamp, List<float> temperatures)>();
        //定义图标X轴元素最大个数
        static int MaxChartSize = 0;
        // 将 linkState 声明为类的成员变量
        private bool linkState = false;
        // 在类的外部定义一个全局变量用于存储当前获取到的温度数据
        private List<float> newDataList = new List<float>();
        private List<float> newDataList2 = new List<float>();
        //配置文件路径
        static string configFilePath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");
        //解析对象JObject
        JObject jsonConfig = JObject.Parse(File.ReadAllText(configFilePath));
        // 创建一个字典来存储每个通道的温度数据
        Dictionary<int, Queue<float>> channelData = new Dictionary<int, Queue<float>>();
        Dictionary<int, Queue<float>> channelData2 = new Dictionary<int, Queue<float>>();
        // 创建列表以存储所有通道的时间点
        Queue<string> dateTimes = new Queue<string>();
        Queue<string> dateTimes2 = new Queue<string>();
        //连接
        static string address = "";
        //起始模块
        StringBuilder addressBuilder = new StringBuilder();
        public Form1()
        {
            InitializeComponent();
            //初始化通道连接属性
            InitializeLoad();
            //加载通道别名
            LoadTextBoxValues();
        }
        private void LoadTextBoxValues()
        {
            tempeartureOtherNameBoxes = new TextBox[] {
            textBox5, textBox6, textBox7, textBox8, textBox12, textBox11, textBox10, textBox9,
            textBox20, textBox19, textBox18, textBox17, textBox28, textBox27, textBox26, textBox25,
            textBox36, textBox35, textBox34, textBox33, textBox44, textBox43, textBox42, textBox41
            };
            tempeartureOtherNameBoxes2 = new TextBox[] {
            textBox56, textBox51, textBox57, textBox58, textBox66, textBox68, textBox73, textBox76,
            textBox87, textBox88, textBox89, textBox90, textBox95, textBox96, textBox97, textBox98,
            textBox94, textBox93, textBox92, textBox91, textBox85, textBox83, textBox81, textBox80
            };
            temperatureCheckBoxes = new CheckBox[]
            {
                checkBox2,checkBox3,checkBox4,checkBox5,checkBox6,checkBox7,checkBox8,checkBox9,
                checkBox17,checkBox16,checkBox15,checkBox14,checkBox13,checkBox12,checkBox11,checkBox10,
                checkBox25,checkBox24,checkBox23,checkBox22,checkBox21,checkBox20,checkBox19,checkBox18
            };
            temperatureCheckBoxes2 = new CheckBox[]
            {
                checkBox49,checkBox48,checkBox47,checkBox46,checkBox45,checkBox44,checkBox43,checkBox42,
                checkBox41,checkBox40,checkBox39,checkBox38,checkBox37,checkBox36,checkBox35,checkBox34,
                checkBox33,checkBox32,checkBox31,checkBox30,checkBox29,checkBox28,checkBox27,checkBox26
            };
            int maxChartSize = (int)jsonConfig["config"]["MaxChartSize"];
            MaxChartSize = (int)maxChartSize;
            try
            {
                if (File.Exists(configFilePath))
                {
                    //加载通道别名
                    JObject channel = (JObject)jsonConfig["config"]["channel1"];
                    JObject channel2 = (JObject)jsonConfig["config"]["channel2"];
                    int index = 0;
                    int index2 = 0;
                    foreach (var kvp in channel)
                    {
                        if (index < tempeartureOtherNameBoxes.Length)
                        {
                            tempeartureOtherNameBoxes[index].Text = kvp.Value.ToString();
                            index++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    foreach (var kvp in channel2)
                    {
                        if (index2 < tempeartureOtherNameBoxes2.Length)
                        {
                            tempeartureOtherNameBoxes2[index2].Text = kvp.Value.ToString();
                            index2++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading configuration: " + ex.Message);
            }
        }

        public const int WM_DEVICE_CHANGE = 0x219;
        public const int DBT_DEVICEARRIVAL = 0x8000;
        public const int DBT_DEVICE_REMOVE_COMPLETE = 0x8004;
        /// <summary>
        /// 检测USB串口的拔插
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_DEVICE_CHANGE) // 捕获USB设备的拔出消息WM_DEVICECHANGE
            {
                switch (m.WParam.ToInt32())
                {
                    case DBT_DEVICE_REMOVE_COMPLETE: // USB拔出 
                        {
                            new Thread(
                                new ThreadStart(
                                    new Action(() =>
                                    {
                                        getPortDeviceName();
                                    })
                                )
                            ).Start();
                        }
                        break;
                    case DBT_DEVICEARRIVAL: // USB插入获取对应串口名称     
                        {
                            new Thread(
                                new ThreadStart(
                                    new Action(() =>
                                    {
                                        getPortDeviceName();
                                    })
                                )
                            ).Start();
                        }
                        break;
                }
            }
            base.WndProc(ref m);
        }

        /// <summary>
        /// 获取串口完整名字（包括驱动名字）
        /// 如果找不到类，需要添加System.Management引用，添加引用->程序集->System.Management
        /// </summary>
        Dictionary<String, String> coms = new Dictionary<String, String>();
        private void getPortDeviceName()
        {

            coms.Clear();
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher
            ("select * from Win32_PnPEntity where Name like '%(COM%'"))
            {
                var hardInfos = searcher.Get();
                foreach (var hardInfo in hardInfos)
                {
                    if (hardInfo.Properties["Name"].Value != null)
                    {
                        string deviceName = hardInfo.Properties["Name"].Value.ToString();
                        int startIndex = deviceName.IndexOf("(");
                        int endIndex = deviceName.IndexOf(")");
                        string key = deviceName.Substring(startIndex + 1, deviceName.Length - startIndex - 2);
                        string name = deviceName.Substring(0, startIndex - 1);
                        Console.WriteLine("key:" + key + ",name:" + name + ",deviceName:" + deviceName);
                        coms.Add(key, name);
                    }
                }

                //创建一个用来更新UI的委托 (主线程更新)
                this.Invoke(
                     new Action(() =>
                     {
                         comboBox1.Items.Clear();
                         foreach (KeyValuePair<string, string> kvp in coms)
                         {
                             comboBox1.Items.Add(kvp.Key);//更新下拉列表中的串口
                         }

                     })
                 );

            }

        }

        private void InitializeLoad()
        {
            string[] ports = SerialPort.GetPortNames();
            if (ports.Length == 0)
            {
                MessageBox.Show("本机没有串口！");
            }
            Array.Sort(ports);
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
            }
            comboBox1.Text = ports[0];
            comboBox2.Text = "19200";
            comboBox3.Text = "1";
            comboBox4.Text = "8";
            // 将所有的温度 TextBox(1) 控件添加到数组中
            temperatureTextBoxes = new TextBox[] {
            textBox1, textBox2, textBox3, textBox4, textBox16, textBox15, textBox14, textBox13,
            textBox24, textBox23, textBox22, textBox21, textBox32, textBox31, textBox30, textBox29,
            textBox40, textBox39, textBox38, textBox37, textBox48, textBox47, textBox46, textBox45
            };
            //将所有的温度 TextBox(2) 控件添加到数组中
            temperatureTextBoxes2 = new TextBox[] {
            textBox52, textBox55 , textBox54 , textBox53 , textBox86 , textBox70 , textBox64 , textBox59 ,
            textBox84, textBox78 , textBox71 , textBox60 , textBox82 , textBox74 , textBox69 , textBox63 ,
            textBox79 , textBox72 , textBox67 , textBox62 , textBox77 , textBox75 , textBox65 , textBox61
            };
            // 手动定义通道号与对应的空列表
            for (int i = 1; i <= 24; i++)
            {
                channelData.Add(i, new Queue<float>());
            }
            for (int i = 1; i <= 24; i++)
            {
                channelData2.Add(i, new Queue<float>());
            }
            chart1.MouseWheel += new MouseEventHandler(_MouseWheel);
        }
        //鼠标滚轮事件
        void _MouseWheel(object sender,MouseEventArgs e)
        {
            if(e.Delta == 120)
            {
                if (chart1.ChartAreas[0].AxisX.ScaleView.Size > 0)
                {
                    chart1.ChartAreas[0].AxisX.ScaleView.Size /= 2;
                }
                else
                {
                    chart1.ChartAreas[0].AxisX.ScaleView.Size = 2;
                }
            }
            else if(e.Delta == -120)
            {
                if (chart1.ChartAreas[0].AxisX.ScaleView.Size > 0)
                {
                    chart1.ChartAreas[0].AxisX.ScaleView.Size *= 2;
                }
                else
                {
                    chart1.ChartAreas[0].AxisX.ScaleView.Size = 2;
                }
            }
        }
        //button连接
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (!linkState)
                {
                    string[] ports = System.IO.Ports.SerialPort.GetPortNames();
                    serialPort.PortName = comboBox1.Text;
                    serialPort.BaudRate = int.Parse(comboBox2.Text); // 波特率
                    serialPort.DataBits = int.Parse(comboBox4.Text); // 数据位
                    if (comboBox3.Text == "1") { serialPort.StopBits = StopBits.One; }
                    else if (comboBox3.Text == "1.5") { serialPort.StopBits = StopBits.OnePointFive; }
                    else if (comboBox3.Text == "2") { serialPort.StopBits = StopBits.Two; }
                    serialPort.Encoding = Encoding.GetEncoding("GB2312"); // 此行非常重要，解决接收中文乱码的问题                                  
                    serialPort.Open(); // 打开串口
                    address = "01";
                    //string msgOrder = "00 03 00 00 00 01 85 DB";
                    //byte[] bytesToSend = base_Conversion.StringToByteArray(msgOrder);
                    //serialPort.Write(bytesToSend, 0, bytesToSend.Length);
                    /*Thread.Sleep(150);
                    byte[] receivedBytes = new byte[serialPort.BytesToRead];
                    serialPort.Read(receivedBytes, 0, receivedBytes.Length);
                    if (receivedBytes.Length == 0)
                    {
                        MessageBox.Show("未找到该模块,请重新连接!");
                        serialPort.Dispose();
                        serialPort.Close();
                        linkState = false;
                    }
                    else
                    {*/
                    /*string hexString = base_Conversion.ByteArrayToString(receivedBytes);//字符串接收
                    byte[] tempByte = base_Conversion.StringToHexBytes(hexString);//字节接收
                    int size = Convert.ToInt32(tempByte[2]);
                    for (int i = 1; i <= size; i++)
                    {
                        address += Convert.ToInt32(tempByte[i + 2]).ToString();
                    }
                    label39.Text = address;*/
                    button2.Text = "断开";
                    linkState = true;
                    checkBox1.Enabled = true;
                    button5.Enabled = true;
                    /*}*/
                }
                else
                {
                    serialPort.Dispose();
                    serialPort.Close();
                    button2.Text = "连接";
                    linkState = false;
                    checkBox1.Enabled = false;
                    button5.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                // 捕获到异常信息，创建一个新的 SerialPort 对象，之前的不能用了。  
                serialPort = new System.IO.Ports.SerialPort();
                // 将异常信息传递给用户。  
                MessageBox.Show(ex.Message);
            }
        }


        //button发送
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort.IsOpen == false)
                {
                    return;
                }
                string msgOrder = txt_Msg.Text;
                byte[] bytesToSend = base_Conversion.StringToByteArray(msgOrder);
                serialPort.Write(bytesToSend, 0, bytesToSend.Length);
                Thread.Sleep(150);
                byte[] receivedBytes = new byte[serialPort.BytesToRead];
                serialPort.Read(receivedBytes, 0, receivedBytes.Length);
                string hexString = base_Conversion.ByteArrayToString(receivedBytes);//字符串接收
                byte[] tempByte = base_Conversion.StringToHexBytes(hexString);//字节接收
                int tempLength = tempByte.Length;
                // 确定截取的起始位置（第六个字节）
                int startIndex = 5;
                // 确定截取的结束位置（倒数第四个字节）
                int endIndex = tempLength - 2;
                // 计算截取的字节个数
                int length = endIndex - startIndex;
                // 创建dataByte数组，长度为截取的字节个数
                byte[] dataByte = new byte[length];
                // 复制tempByte数组中的指定范围到dataByte数组中
                Array.Copy(tempByte, startIndex, dataByte, 0, length);
                // 定义每组字节的长度
                int groupSize = 4;
                // 计算总共有多少组
                int numberOfGroups = length / groupSize;
                // 创建一个列表用于存储每组字节
                List<byte[]> bytes = new List<byte[]>();

                // 循环遍历每一组字节
                for (int i1 = 0; i1 < numberOfGroups; i1++)
                {
                    // 计算当前组的起始索引time3
                    int groupStartIndex = i1 * groupSize;
                    // 创建一个长度为groupSize的数组用于存储当前组的字节
                    byte[] groupBytes = new byte[groupSize];
                    // 复制当前组的字节到groupBytes数组中
                    Array.Copy(dataByte, groupStartIndex, groupBytes, 0, groupSize);
                    // 将当前组的字节数组添加到bytes列表中
                    bytes.Add(groupBytes);
                }
                int index = 0;
                foreach (var byte1 in bytes)
                {
                    // 反转字节数组的顺序
                    Array.Reverse(byte1);
                    //Debug.WriteLine(BitConverter.ToSingle(byte1, 0));
                    if (BitConverter.ToSingle(byte1, 0) < 3000)
                    {
                        temperatureTextBoxes[index].Text = (BitConverter.ToSingle(byte1, 0)).ToString();
                    }
                    else
                    {
                        temperatureTextBoxes[index].Text = "0";
                    }
                    index++;
                }

                this.Invoke((MethodInvoker)(() => txt_Received.Text += hexString));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void SaveTemperatureDataToExcel(Queue<(DateTime timeStamp, List<float> temperatures)> temperatureData, Queue<(DateTime timeStamp, List<float> temperatures)> temperatureData2)
        {
            try
            {
                // 创建一个新的Excel工作簿
                HSSFWorkbook workbook = new HSSFWorkbook();

                // 添加一个工作表
                ISheet sheet = workbook.CreateSheet("TemperatureData");
                ISheet sheet2 = workbook.CreateSheet("TemperatureData2");

                // 行数计数器
                int currentRowNumber = 0;
                int currentRowNumber2 = 0;

                // 添加表头
                IRow headerRow = sheet.CreateRow(currentRowNumber);
                IRow headerRow2 = sheet2.CreateRow(currentRowNumber2);
                headerRow.CreateCell(0).SetCellValue("时间");
                headerRow2.CreateCell(0).SetCellValue("时间");
                // 设置时间戳列的宽度
                sheet.SetColumnWidth(0, 20 * 256); // 20个字符的宽度，单位为 1/256 字符
                sheet2.SetColumnWidth(0, 20 * 256);
                //加载通道别名
                JObject channel = (JObject)jsonConfig["config"]["channel1"];
                JObject channel2 = (JObject)jsonConfig["config"]["channel2"];
                List<string> channelOtherName = new List<string>();
                List<string> channelOtherName2 = new List<string>();

                foreach (var kvp in channel)
                {
                    channelOtherName.Add(kvp.Value.ToString());
                }
                foreach (var kvp in channel2)
                {
                    channelOtherName2.Add(kvp.Value.ToString());
                }

                for (int i = 1; i <= 24; i++)
                {
                    headerRow.CreateCell(i).SetCellValue("通道" + i + "(" + channelOtherName[i - 1].ToString() + ")");
                    // 设置通道列的宽度
                    sheet.SetColumnWidth(i, 15 * 256); // 15个字符的宽度，单位为 1/256 字符
                }
                for (int i = 1; i <= 24; i++)
                {
                    headerRow2.CreateCell(i).SetCellValue("通道" + i + "(" + channelOtherName2[i - 1].ToString() + ")");
                    // 设置通道列的宽度
                    sheet2.SetColumnWidth(i, 15 * 256); // 15个字符的宽度，单位为 1/256 字符
                }
                currentRowNumber++;
                currentRowNumber2++;
                // 写入数据到工作表中
                foreach (var dataPoint in temperatureData)
                {
                    var timeStamp = dataPoint.timeStamp;
                    var temperatures = dataPoint.temperatures;

                    // 检查是否有数据
                    if (temperatures.Any())
                    {
                        // 创建行
                        IRow row = sheet.CreateRow(currentRowNumber);

                        // 在第一列插入时间戳
                        row.CreateCell(0).SetCellValue(timeStamp.ToString("yyyy-MM-dd HH:mm:ss:fff"));

                        // 写入温度数据到后续列
                        int currentColumn = 1; // 从第二列开始写入温度数据
                        foreach (var temperature in temperatures)
                        {
                            // 设置单元格值
                            row.CreateCell(currentColumn).SetCellValue((double)temperature);

                            currentColumn++;
                        }

                        currentRowNumber++;
                    }
                }
                foreach (var dataPoint in temperatureData2)
                {
                    var timeStamp = dataPoint.timeStamp;
                    var temperatures = dataPoint.temperatures;

                    // 检查是否有数据
                    if (temperatures.Any())
                    {
                        // 创建行
                        IRow row = sheet2.CreateRow(currentRowNumber2);

                        // 在第一列插入时间戳
                        row.CreateCell(0).SetCellValue(timeStamp.ToString("yyyy-MM-dd HH:mm:ss:fff"));

                        // 写入温度数据到后续列
                        int currentColumn = 1; // 从第二列开始写入温度数据
                        foreach (var temperature in temperatures)
                        {
                            // 设置单元格值
                            row.CreateCell(currentColumn).SetCellValue((double)temperature);

                            currentColumn++;
                        }

                        currentRowNumber2++;
                    }
                }

                // 保存工作簿到文件
                string filePath = textBox50.Text + "\\TemperatureData" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".csv"; // 你想要保存的自定义路径
                using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                    MessageBox.Show("Excel 文件已保存到：" + filePath, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存Excel文件时发生错误：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string savePath = textBox50.Text; // 获取保存路径

                // 检查路径是否存在
                if (!Directory.Exists(savePath))
                {
                    MessageBox.Show("指定路径不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return; // 中断执行
                }

                if (button4.Text == "已停止")
                {
                    button4.Text = "记录中...";
                }
                else if (button4.Text == "记录中...")
                {
                    SaveTemperatureDataToExcel(temperatureData, temperatureData2);
                    button4.Text = "已停止";
                    temperatureData.Clear();
                    temperatureData2.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

            if(checkBox1.Checked == true)
            {
                 timer1.Start();
                 timer2.Start();
            }
            else
            {
                timer1.Stop();
                timer2.Stop();
            }
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if ( !illegal.IsValidHexadecimal(base_Conversion.DecimalToHexadecimal(textBox99.Text)) || !illegal.IsValidHexadecimal(base_Conversion.DecimalToHexadecimal(textBox100.Text)))
                {
                    throw new ArgumentException("输入格式错误");
                }
                else
                {
                    string temp1 = base_Conversion.DecimalToHexadecimal(textBox99.Text) + " 03 00 00 00 31";
                    string msgOrder1 = temp1 + " " + base_Conversion.ByteArrayToString(crc16.ToModbus(base_Conversion.StringToHexBytes(temp1), 6));
                    dateTimes.Enqueue(DateTime.Now.ToString("HH:mm:ss.fff"));
                    HandleData1(msgOrder1);
                    //string temp11 = base_Conversion.DecimalToHexadecimal(textBox99.Text) + "04 00 3F 00 18";

                    for (int i = 0; i < channelData.Count; i++)
                    {
                        if (dateTimes.Count != channelData[i + 1].Count && dateTimes.Count >= 0)
                        {
                            dateTimes.Dequeue();
                        }
                        else
                        {
                            chart1.Series[i].Points.DataBindXY(dateTimes, channelData[i + 1]);
                        }
                        if (dateTimes.Count >= MaxChartSize && channelData[i + 1].Count >= MaxChartSize)
                        {
                            dateTimes.Dequeue();
                            foreach (var channel in channelData)
                            {
                                var key = channel.Key;
                                channelData[key].Dequeue();
                            }
                        }
                    }
                    string temp2 = base_Conversion.DecimalToHexadecimal(textBox100.Text) + " 03 00 00 00 31";
                    string msgOrder2 = temp2 + " " + base_Conversion.ByteArrayToString(crc16.ToModbus(base_Conversion.StringToHexBytes(temp2), 6));
                    dateTimes2.Enqueue(DateTime.Now.ToString("HH:mm:ss.fff"));
                    HandleData2(msgOrder2);
                    for (int i = 0; i < channelData2.Count; i++)
                    {
                        if(dateTimes2.Count != channelData2[i + 1].Count && dateTimes2.Count >= 0)
                        {
                            dateTimes2.Dequeue();
                        }
                        else
                        {
                            chart1.Series[i + temperatureCheckBoxes.Length].Points.DataBindXY(dateTimes2, channelData2[i + 1]);
                        }
                        if (dateTimes2.Count >= MaxChartSize && channelData2[i + 1].Count >= MaxChartSize)
                        {
                            dateTimes2.Dequeue();
                            foreach (var channel in channelData2)
                            {
                                var key = channel.Key;
                                channelData2[key].Dequeue();
                            }
                        }
                    }

                    if (button4.Text == "记录中...")
                    {
                        this.Invoke((MethodInvoker)(() =>
                        {
                            if(dateTimes.Count > 0)
                            {
                                temperatureData.Enqueue((DateTime.Parse(dateTimes.Last()), new List<float>(newDataList)));
                            }
                            if(dateTimes2.Count > 0)
                            {
                                temperatureData2.Enqueue((DateTime.Parse(dateTimes2.Last()), new List<float>(newDataList2)));
                            }
                        }));
                    }

                    newDataList.Clear();
                    newDataList2.Clear();
                }
            }
            catch (Exception ex)
            {
                button2.Text = "连接";
                serialPort.Close();
                linkState = false;
                checkBox1.Checked = false;
                timer1.Stop();
                MessageBox.Show(ex.Message);
            }
        }

        public void HandleData1(string msgOrder)
        {
            byte[] bytesToSend = base_Conversion.StringToByteArray(msgOrder);
            serialPort.Write(bytesToSend, 0, bytesToSend.Length);
            Thread.Sleep(150);
            byte[] receivedBytes = new byte[serialPort.BytesToRead];
            if(receivedBytes.Length == 0)
            {
                for(int i = 0; i < temperatureTextBoxes.Length; i++)
                {
                    temperatureTextBoxes[i].Text = "";
                    temperatureTextBoxes[i].BackColor = Color.White;
                }
                for (int i = 0; i < temperatureCheckBoxes.Length; i++)
                {
                    dateTimes.Clear();
                    foreach (var queue in channelData.Values)
                    {
                        queue.Clear();
                    }
                    chart1.Series[i].Points.Clear();
                }
                return;
            }
            serialPort.Read(receivedBytes, 0, receivedBytes.Length);
            string hexString = base_Conversion.ByteArrayToString(receivedBytes);
            byte[] tempByte = base_Conversion.StringToHexBytes(hexString);
            int tempLength = tempByte.Length;

            int startIndex = 5;
            int endIndex = tempLength - 2;
            int length = endIndex - startIndex;
            byte[] dataByte = new byte[length];
            Array.Copy(tempByte, startIndex, dataByte, 0, length);

            int groupSize = 4;
            int numberOfGroups = length / groupSize;
            Debug.WriteLine("群数"+numberOfGroups);
            List<byte[]> bytes = new List<byte[]>();
            List<string> timestamps = new List<string>();

            for (int i = 0; i < numberOfGroups; i++)
            {
                int groupStartIndex = i * groupSize;
                byte[] groupBytes = new byte[groupSize];
                Array.Copy(dataByte, groupStartIndex, groupBytes, 0, groupSize);
                bytes.Add(groupBytes);
            }

            JObject forewarning1 = (JObject)jsonConfig["config"]["forewarning1"];
            List<float> lowerLimit = new List<float>();
            List<float> highLimit = new List<float>();
            int index = 0;

            foreach(var byte2 in bytes)
            {
                foreach(var byte1 in byte2)
                {
                    Debug.WriteLine(byte1.ToString());
                }
                Debug.WriteLine("-------------------");
            }

            foreach (var byte1 in bytes)
            {
                Array.Reverse(byte1);
                float temperature = BitConverter.ToSingle(byte1, 0);
                temperatureTextBoxes[index].Text = temperature.ToString();
                newDataList.Add(temperature);

                channelData[index + 1].Enqueue(temperature);
                timestamps.Add(DateTime.Now.ToString());

                for (int i = 1; i <= 24; i++)
                {
                    lowerLimit.Add((float)forewarning1["low" + i.ToString("00")]);
                    highLimit.Add((float)forewarning1["high" + i.ToString("00")]);
                }

                if (temperature < lowerLimit[index] || temperature > highLimit[index])
                {
                    temperatureTextBoxes[index].BackColor = Color.DarkOrange;
                }
                else
                {
                    temperatureTextBoxes[index].BackColor = Color.White;
                }

                index++;
            }
        }

        public void HandleData2(string msgOrder)
        {
            byte[] bytesToSend = base_Conversion.StringToByteArray(msgOrder);
            serialPort.Write(bytesToSend, 0, bytesToSend.Length);
            Thread.Sleep(150);
            byte[] receivedBytes = new byte[serialPort.BytesToRead];
            if (receivedBytes.Length == 0)
            {
                for (int i = 0; i < temperatureTextBoxes2.Length; i++)
                {
                    temperatureTextBoxes2[i].Text = "";
                    temperatureTextBoxes2[i].BackColor = Color.White;
                }
                for (int i = 0; i < temperatureCheckBoxes2.Length; i++)
                {
                    dateTimes2.Clear();
                    foreach (var queue in channelData2.Values)
                    {
                        queue.Clear();
                    }
                    chart1.Series[i + temperatureCheckBoxes.Length].Points.Clear();
                }
                return;
            }
            serialPort.Read(receivedBytes, 0, receivedBytes.Length);
            string hexString = base_Conversion.ByteArrayToString(receivedBytes);
            byte[] tempByte = base_Conversion.StringToHexBytes(hexString);
            int tempLength = tempByte.Length;

            int startIndex = 5;
            int endIndex = tempLength - 2;
            int length = endIndex - startIndex;
            byte[] dataByte = new byte[length];
            Array.Copy(tempByte, startIndex, dataByte, 0, length);

            int groupSize = 4;
            int numberOfGroups = length / groupSize;
            List<byte[]> bytes = new List<byte[]>();
            List<string> timestamps = new List<string>();

            for (int i = 0; i < numberOfGroups; i++)
            {
                int groupStartIndex = i * groupSize;
                byte[] groupBytes = new byte[groupSize];
                Array.Copy(dataByte, groupStartIndex, groupBytes, 0, groupSize);
                bytes.Add(groupBytes);
            }

            JObject forewarning2 = (JObject)jsonConfig["config"]["forewarning2"];
            List<float> lowerLimit = new List<float>();
            List<float> highLimit = new List<float>();
            int index = 0;

            foreach (var byte1 in bytes)
            {
                Array.Reverse(byte1);
                float temperature = BitConverter.ToSingle(byte1, 0);
                temperatureTextBoxes2[index].Text = temperature.ToString();
                newDataList2.Add(temperature);

                channelData2[index + 1].Enqueue(temperature);
                timestamps.Add(DateTime.Now.ToString());

                for (int i = 1; i <= 24; i++)
                {
                    lowerLimit.Add((float)forewarning2["low" + i.ToString("00")]);
                    highLimit.Add((float)forewarning2["high" + i.ToString("00")]);
                }

                if (temperature < lowerLimit[index] || temperature > highLimit[index])
                {
                    temperatureTextBoxes2[index].BackColor = Color.DarkOrange;
                }
                else
                {
                    temperatureTextBoxes2[index].BackColor = Color.White;
                }

                index++;
            }
        }



        //选择路径
        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                DirectoryInfo theFolder = new DirectoryInfo(foldPath);
                string dirInfo = theFolder.FullName;
                textBox50.Text = dirInfo;
            }
        }
        //关闭Form1
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (serialPort != null && serialPort.IsOpen)
                {
                    serialPort.Close();
                    serialPort.Dispose();
                }
            }
            catch (Exception ex)
            {
                //将异常信息传递给用户。  
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < temperatureCheckBoxes.Length; i++)
            {
                if (!temperatureCheckBoxes[i].Checked)
                {
                    chart1.Series[i].Enabled = false;
                }
                else
                {
                    chart1.Series[i].Enabled = true;
                }
            }
            for (int i = 0; i < temperatureCheckBoxes2.Length; i++)
            {
                if (!temperatureCheckBoxes2[i].Checked)
                {
                    chart1.Series[i + temperatureCheckBoxes.Length].Enabled = false;
                }
                else
                {
                    chart1.Series[i + temperatureCheckBoxes.Length].Enabled = true;
                }
            }
        }

        //修改地址
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string address = base_Conversion.DecimalToHexadecimal(textBox49.Text);
                string temp = addressBuilder+ " 06 00 00 00 " + address;
                string msgOrder = addressBuilder + " 06 00 00 00 " + address + " " + base_Conversion.ByteArrayToString(crc16.ToModbus(base_Conversion.StringToHexBytes(temp), 6));
                
                byte[] bytesToSend = base_Conversion.StringToByteArray(msgOrder);
                serialPort.Write(bytesToSend,0,bytesToSend.Length);
                // 等待一段时间以确保数据接收完整
                Thread.Sleep(150);
                byte[] receivedBytes = new byte[serialPort.BytesToRead];
                serialPort.Read(receivedBytes, 0, receivedBytes.Length);
                string hexString = base_Conversion.ByteArrayToString(receivedBytes);//字符串接收
                byte[] tempByte = base_Conversion.StringToHexBytes(hexString);//字节接收
                txt_Received.Text += hexString;
                addressBuilder.Clear();
                addressBuilder.Append(address);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        //搜索地址
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                string msgOrder = "00 03 00 00 00 01 85 DB";
                byte[] bytesToSend = base_Conversion.StringToByteArray(msgOrder);
                serialPort.Write(bytesToSend, 0, bytesToSend.Length);
                Thread.Sleep(150);
                byte[] receivedBytes = new byte[serialPort.BytesToRead];
                serialPort.Read(receivedBytes, 0, receivedBytes.Length);
                if (receivedBytes.Length == 0)
                {
                    MessageBox.Show("未找到该模块,请重新连接!");
                    serialPort.Dispose();
                    serialPort.Close();
                    linkState = false;
                }
                else
                {
                    int size = receivedBytes[2];
                    for (int i = 0; i < size; i++)
                    {
                        addressBuilder.Append(Convert.ToInt32(receivedBytes[i + 3]).ToString());
                    }
                    textBox49.Text = addressBuilder.ToString();
                    groupBox2.Text = "通道温度检测数据" + addressBuilder.ToString();
                    addressBuilder.Clear();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

    }
}
