using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSenderConsole
{
    public static class Log
    {
        //public static string logPath{ get; set; }
        /// <summary>
        /// 写入指定日志文件
        /// </summary>
        /// <param name="logContent"></param>
        /// <param name="logPath"></param>
        public static void WriteLog(string logContent, string logPath)
        {
            FileStream fs = null;
            string filePath = logPath;
            //将待写的入数据从字符串转换为字节数组
            Encoding encoder = Encoding.UTF8;
            byte[] bytes = encoder.GetBytes(logContent + "\n\r");
            try
            {
                fs = File.OpenWrite(filePath);
                //设定书写的开始位置为文件的末尾
                fs.Position = fs.Length;
                //将待写入内容追加到文件末尾
                fs.Write(bytes, 0, bytes.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine("文件打开失败{0}", ex.ToString());
            }
            finally
            {
                fs.Close();
            }
        }
    }
}
