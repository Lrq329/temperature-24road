using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test11.Tools
{
    public class Base_conversion
    {
        public byte[] StringToByteArray(string hex)
        {
            // 去除字符串中的空格
            hex = hex.Replace(" ", "");
            int NumberChars = hex.Length / 2;
            byte[] bytes = new byte[NumberChars];
            using (var sr = new StringReader(hex))
            {
                for (int i = 0; i < NumberChars; i++)
                    bytes[i] = Convert.ToByte(new string(new char[2] { (char)sr.Read(), (char)sr.Read() }), 16);
            }
            return bytes;
        }

        public string ByteArrayToString(byte[] bytes)
        {
            StringBuilder hex = new StringBuilder(bytes.Length * 2);
            foreach (byte b in bytes)
                hex.AppendFormat("{0:X2} ", b);
            return hex.ToString();
        }

        public string DecimalToHexadecimal(string decimalString)
        {
            // 将十进制字符串转换为整数
            int decimalValue = int.Parse(decimalString);

            // 将整数转换为十六进制字符串
            string hexString = decimalValue.ToString("X2");

            // 返回十六进制字符串
            return hexString;
        }

        public byte[] StringToHexBytes(string input)
        {
            // 去除字符串中的空格并将其转换为大写形式
            input = input.Replace(" ", "").ToUpper();

            // 确保输入字符串的长度为偶数
            if (input.Length % 2 != 0)
            {
                throw new ArgumentException("输入格式错误");
            }

            // 创建字节数组以保存结果
            byte[] bytes = new byte[input.Length / 2];

            // 逐个字节将字符串转换为字节数组
            for (int i = 0; i < input.Length; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(input.Substring(i, 2), 16);
            }

            return bytes;
        }

    }
}
