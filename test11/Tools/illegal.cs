using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test11.Tools
{
    public class illegal
    {
        public bool IsValidHexadecimal(string hexString)
        {
            // 检查字符串是否为空或长度为奇数（十六进制字符串的长度应为偶数）
            if (string.IsNullOrEmpty(hexString) || hexString.Length % 2 != 0)
            {
                return false;
            }

            // 检查字符串中的每个字符是否都是十六进制数字
            foreach (char c in hexString)
            {
                if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')))
                {
                    return false;
                }
            }

            // 字符串合法
            return true;
        }


    }
}
