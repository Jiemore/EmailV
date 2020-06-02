using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailSenderConsole
{
    public class Sender
    {
        //账号
        public string MyEmail { get; set; }
        //密码
        public string MyPwd { get; set; }
        //掉线次数
        public int Down { get; set; }
        //发送总数
        public int SendTotal { get; set; }
        //失败次数
        public int FaultTotal { get; set; }
        //成功次数
        public int SuccessTotal { get; set; }

        public Sender() {
            Down = 0;
            SendTotal = 0;
            FaultTotal = 0;
            SuccessTotal = 0;
        }
    }
}
