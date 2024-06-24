using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test11.Tools
{
    public class CRC16MODBUS
    {
        public class CRC16
        {
            /// <summary>
            /// CRC16_Modbus效验
            /// </summary>
            /// <param name="byteData">要进行计算的字节数组</param>
            /// <param name="byteLength">长度</param>
            /// <returns>计算后的数组</returns>
            public byte[] ToModbus(byte[] byteData, int byteLength)
            {
                byte[] CRC = new byte[2];

                ushort wCrc = 0xFFFF;
                for (int i = 0; i < byteLength; i++)
                {
                    wCrc ^= Convert.ToUInt16(byteData[i]);
                    for (int j = 0; j < 8; j++)
                    {
                        if ((wCrc & 0x0001) == 1)
                        {
                            wCrc >>= 1;
                            wCrc ^= 0xA001;//异或多项式
                        }
                        else
                        {
                            wCrc >>= 1;
                        }
                    }
                }
                CRC[1] = (byte)((wCrc & 0xFF00) >> 8);//高位在后
                CRC[0] = (byte)(wCrc & 0x00FF);       //低位在前
                return CRC;

            }
        }
    }
}
