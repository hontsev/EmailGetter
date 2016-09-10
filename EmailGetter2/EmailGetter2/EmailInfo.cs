using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace EmailGetter2
{
    public class EmailInfo
    {
        public string From;
        public string To;
        public string Cc;
        public string Date;
        public string Subject;
        public string Content;
        public int no;
        public EmailInfo()
        {
            From = "";
            To = "";
            Cc = "";
            Subject = "";
            Content = "";
            Date = "";
            no = 0;
        }
    }
}
