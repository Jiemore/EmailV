using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EmailSenderConsole
{
    class Program
    {
        static ImportData data = new ImportData();
        static string SendFile;
        static string RecvFile;
        static string LogFile;

        //日志：发送记录
        static string sendRecord;
        static string EmailServer;
        static int Port;

        static string Content = "大量商家入驻抖音直播卖货，急需点赞手！\n添加代理ID：436485431";
        static string Subject = "抖音点赞*兼职*福利！";

        static void Main(string[] args)
        {
            Console.WriteLine("Setting Loading...");
            //读取配置
            ReadConfig();
            Console.WriteLine("Settings Loading Done.");

            Console.WriteLine("Senders Initialization...");
            if (!data.ImportSender(SendFile)) {
                Console.WriteLine("发送人初始化...失败！\n输入【回车】退出程序...");
                Console.ReadLine();
                return;
            }
            Console.WriteLine("Senders Loading Done.");
            Console.WriteLine("Receiver Initialization...");
            if (!data.ImportReceiver(RecvFile))
            {
                Console.WriteLine("收件人初始化...失败！\n输入【回车】退出程序...");
                Console.ReadLine();
                return;
            }
            Console.WriteLine("Receiver Loading Done.");
            Console.WriteLine($">>>>>>>>>>>>>>>>>>>>>>>{DateTime.Now.ToString()}>>>>>>>>>>>>>>>>>>>>>>>>>");

            string MyPsd;//这是邮箱密码
            string MyEmail;//这是邮箱账号

            while (true)
            {
                foreach (var item in data.Senders)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        MyEmail = item.MyEmail;
                        MyPsd = item.MyPwd;

                        Console.WriteLine("Email Message Initialization....");
                        MailMessage MMsg = new MailMessage();
                        MMsg.Subject = Subject;//这是邮件主题
                        MMsg.From = new MailAddress(MyEmail);

                        if (data.EmailAccount.Count < 1)
                        {

                            Console.WriteLine("全部收件人已发送完毕！\n");
                            Log.WriteLog($"\r\n[Send Done]日期{DateTime.Now}:全部收件人已发送完毕！\n", LogFile);
                            Console.ReadLine();
                            return;
                        }

                        //收件人
                        string rec = "";
                        for (int i = 0; i < 20; i++)
                        {
                            rec = data.EmailAccount.Pop().ToString().Trim();
                            MMsg.To.Add(new MailAddress(rec));//这是邮件地址
                        }

                        Console.WriteLine("添加发送人：" + rec);

                        MMsg.IsBodyHtml = true;//这里启用IsBodyHtml是为了支持内容中的Html。
                        MMsg.BodyEncoding = Encoding.UTF8;//将正文的编码形式设置为UTF8。
                        MMsg.Body = Content;//这是邮件内容
                        object userState = MMsg;
                        Console.WriteLine("邮件信息初始化完毕....");

                        Console.WriteLine("邮件服务器初始化中...");
                        SmtpClient SClient = new SmtpClient();
                        SClient.Host = EmailServer;//"smtp.163.com";//smtp地址
                        SClient.Port = Port;//25;//smtp端口
                        SClient.EnableSsl = true;//因为使用了SSL（安全套接字层）加密链接所以这里的EnableSsl必须设置为true。
                                                 //SClient.Credentials = new NetworkCredential(MyEmail, MyPsd);
                        Console.WriteLine("邮件服务器初始化完毕...");
                        Console.WriteLine(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
                        Console.WriteLine("小单杠（125）开始发射...");
                        try
                        {
                            SClient.SendCompleted += new SendCompletedEventHandler(SmtpClient_OnCompleted);
                            SClient.SendAsync(MMsg, userState);
                            Console.WriteLine(MyEmail + ":邮件已经发送成功!");
                            item.SuccessTotal++;

                            //写入发送日志
                            Log.WriteLog($"{item.MyEmail} To {rec}\r\n", sendRecord);

                        }
                        catch (Exception err)
                        {
                            item.FaultTotal++;
                            //将发送失败的用户账户，重新放入发送堆
                            data.EmailAccount.Push(rec);
                            Console.WriteLine(err.Message, "错误提示!");
                            Console.WriteLine("账号切换...");
                            Log.WriteLog($"{item.MyEmail} 发送错误：{err.Message}\r\n {item.MyEmail} To {rec}\r\n", sendRecord);
                            //换号
                            item.Down++;
                            continue;
                        }
                        finally
                        {
                            item.SendTotal++;
                        }
                    }
                }
                //监测所有账户是否全部挂机
                if (!Isalldown(50))
                    break;
            }
            Console.WriteLine("账号已全部挂机！");
            Console.ReadLine();
        }

        /// <summary>
        /// 允许账号挂机次数
        /// </summary>
        /// <param name="down"></param>
        /// <returns></returns>
        static bool Isalldown(int down) {
            foreach (var item in data.Senders)
            {
                //账号已超过挂机次数，写入日志记录
                if (item.Down >= down)
                {
                    string log = $"发送者：{item.MyEmail} >> 发送总数：{item.SendTotal}\r\n" +
                                 $"挂机次数：{item.Down}>> 发送失败总数：{item.FaultTotal}\r\n" +
                                 $"成功次数：{item.SuccessTotal} \r\n"+
                                 "================================================================\r\n";
                    Log.WriteLog(log,LogFile);
                    data.Senders.Remove(item);
                    break;
                }
            }
            return data.Senders.Count > 1 ? true : false;
        }
        /// <summary>
        /// 读取配置
        /// </summary>
        static void ReadConfig() {
            //配置文件路径
            string settingPath = AppDomain.CurrentDomain.BaseDirectory + "settings.json";

            //文件不存在则创建配置文件
            if (!File.Exists(settingPath))
            {
                WriteSettings();
            }
            else
            {
                var settingString = File.ReadAllText(settingPath);
                dynamic settings = JsonConvert.DeserializeObject<dynamic>(settingString);
                try
                {
                    SendFile = settings["SendFile"];
                    RecvFile = settings["RecvFile"];
                    LogFile = settings["log"];
                    sendRecord = settings["sendRecord"];
                    EmailServer = settings["EmailServer"];
                    Port = Convert.ToInt32(settings["Port"]);
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
        /// <summary>
        /// 创建配置
        /// </summary>
        static void WriteSettings()
        {
            var settingPath = AppDomain.CurrentDomain.BaseDirectory + "settings.json";
            var settings = new Dictionary<string, object>
            {
                { "SendFile",""},
                { "RecvFile",""},
                { "log",""},
                { "sendRecord",""},
                { "EmailServer",""},
                { "Port",""}
            };
            File.WriteAllText(settingPath, JsonConvert.SerializeObject(settings, Formatting.Indented));
        }
        /// <summary>
        /// 写入发送日志
        /// </summary>
        /// <param name="logContent"></param>

        /// <summary>
        /// 邮件异步发送处理结果
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void SmtpClient_OnCompleted(object sender,AsyncCompletedEventArgs e) {
            string strlog = string.Empty;
            MailMessage mail = (MailMessage)e.UserState;
            string subject = mail.Subject;

            if (e.Cancelled)
            {
                //邮件已取消发送
                foreach (var item in mail.To)
                {
                    strlog = $"\r\n[False]From{mail.From} To {item}: Send canceled for mail with subject [{subject}].\r\n";
                    Log.WriteLog(strlog, sendRecord);
                    Console.WriteLine(strlog);
                }
            }
            if (e.Error != null)
            {
                //邮件发送异常
                foreach (var item in mail.To)
                {
                    strlog = $"\r\n[False]From{mail.From} To {item}: Erro!\r\n Error Message : [{e.Error.ToString()}].\r\n";
                    Log.WriteLog(strlog, sendRecord);
                    Console.WriteLine(strlog);
                }
            }
            else
            {
                //已发送
                foreach (var item in mail.To)
                {
                    strlog = $"\r\n[True]From{mail.From} To {item}: Suceess!\r\n";
                    Log.WriteLog(strlog, sendRecord);
                    Console.WriteLine(strlog);
                }
            }
        }
    }
}
