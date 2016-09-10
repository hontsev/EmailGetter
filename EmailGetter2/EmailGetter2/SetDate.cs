using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace EmailGetter2
{
    public partial class SetDate : Form
    {
        public int dateNo = 1;
        public DateTime dt;
        public SetDate()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dt = monthCalendar1.SelectionEnd;
            System.Diagnostics.Debug.Print(dt.ToString());
            this.Owner.Enabled = true;
            if (dateNo == 1)
            {
                ((Form1)(this.Owner)).selectDate1 = new DateTime(monthCalendar1.SelectionStart.Year, monthCalendar1.SelectionStart.Month, monthCalendar1.SelectionStart.Day, 0, 0, 0);
            }
            else
            {
                ((Form1)(this.Owner)).selectDate2 = new DateTime(monthCalendar1.SelectionStart.Year, monthCalendar1.SelectionStart.Month, monthCalendar1.SelectionStart.Day, 23, 59, 59);
            }
            ((Form1)(this.Owner)).updateDates();
            this.Hide();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Owner.Enabled = true;
            this.Hide();
        }
    }
}
