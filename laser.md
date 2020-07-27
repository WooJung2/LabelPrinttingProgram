using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;

namespace LabelPrinttingSystem
{
    public partial class Main_Laser : DevExpress.XtraReports.UI.XtraReport
    {
        public Main_Laser()
        {
            InitializeComponent();
        }

        private int Num;
        public int PageNum
        {
            get { return this.Num; }
            set { this.Num = value; }
        }
        
        private void xrTableCell13_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            xrTableCell13.Text = PageNum.ToString();
            PageNum++;
        }
    }
}
