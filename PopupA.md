using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Common;
using System.IO.Ports;
using Microsoft.Win32;

namespace LabelPrinttingSystem
{
    public delegate void DataGetEventHandler(string PrintType);
    public partial class RegistryForm : DevExpress.XtraEditors.XtraForm
    {
        public DataGetEventHandler dataSendEvent;

        Common.DataService.DataService _service = new Common.DataService.DataService();
        Common.DataService.DataCommand _cmm;
        SerialPort sp = new SerialPort();

        bool ChkIp = false;

        public RegistryForm()
        {
            InitializeComponent();
        }

        private void RegistryForm_Load(object sender, EventArgs e)
        {
            cboPrtInfo.Properties.Items.Add("TCP/IP");
            cboPrtInfo.Properties.Items.Add("Serial");
            cboPrtInfo.EditValue = "TCP/IP";

            searchPortList();
            Setting();
        }

        private void btn_Regit_Click(object sender, EventArgs e)
        {
            _cmm = new Common.DataService.DataCommand(_service);

            string strTable = "PrinttingSetting";

            try
            {
                _service.OpenDB();

                if (_service.IsConn())
                {
                    if (cboPrtInfo.Text == "TCP/IP")
                    {
                        ChkIp = IsAvailableIP(txt_Ip.Text);

                        if (ChkIp)
                        {
                            //���� ������ Ȯ��
                            Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                            Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                            StringBuilder strQuery = new StringBuilder();
                            strQuery.Append("Select PrintValue From PrinttingSetting Where PrintType = '" + cboPrtInfo.Text + "';");

                            DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToStr());
                            strQuery.Clear();
                            DataTable dtHistory = dsResult.Tables[0];
                            bool check = dtHistory.Rows.Count > 0 ? true : false;

                            if (check)
                            {
                                Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                                Dictionary<string, object> parameterTarget = new Dictionary<string, object>();

                                //Set ���� - IP ������Ʈ
                                parameterTarget.Add("PrintValue", txt_Ip.Text);

                                //WHERE ����
                                parameterWhere.Add("PrintType", cboPrtInfo.Text);

                                int Upresult = _cmm.Update(strTable, parameterWhere, parameterTarget);
                                
                                if (Upresult == -1)
                                {
                                    CommonHelper.OpenMessageBox("DB�� ������ �ֽ��ϴ�.", MessageButtonType.OK, MessageIconType.Error);
                                }
                            }
                            else
                            {
                                Dictionary<int, object> parameter = new Dictionary<int, object>();
                                parameter.Add(0, cboPrtInfo.Text);
                                parameter.Add(1, txt_Ip.Text);

                                int result = _cmm.Insert(strTable, parameter);

                                if (result == -1)
                                {
                                    CommonHelper.OpenMessageBox("������ �����Ͽ����ϴ�.", MessageButtonType.OK, MessageIconType.Error);
                                    return;
                                }
                            }

                            CommonHelper.OpenMessageBox("IP�� �����Ǿ����ϴ�.", MessageButtonType.OK, MessageIconType.Information);
                            CommonConnect.PrintType = cboPrtInfo.Text;
                            CommonConnect.PrintIP = txt_Ip.Text;
                            dataSendEvent(cboPrtInfo.Text);
                        }
                        else
                        {
                            CommonHelper.OpenMessageBox("IP�� ��Ȯ�ϰ� �Է����ֽñ� �ٶ��ϴ�.                                  (�� : 000.000.000.000)", MessageButtonType.OK, MessageIconType.Error);
                            return;
                        }

                        //RegistryKey reg = Registry.LocalMachine;

                        //reg = reg.CreateSubKey("SOFTWARE\\TEST");
                        //reg.SetValue("ServerIp", Ip, RegistryValueKind.String);

                        //CommonConnect.PrintType = "TCPIP";
                    }
                    else
                    {
                        if (cboPort.Text == "")
                        {
                            CommonHelper.OpenMessageBox("����Ǿ��ִ� ��Ʈ�� �����ϴ�.", MessageButtonType.OK, MessageIconType.Error);
                            return;
                        }
                        //���� ������ Ȯ��
                        Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                        Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                        StringBuilder strQuery = new StringBuilder();
                        strQuery.Append("Select PrintValue From PrinttingSetting Where PrintType = '" + cboPrtInfo.Text + "';");

                        DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToStr());
                        strQuery.Clear();
                        DataTable dtHistory = dsResult.Tables[0];
                        bool check = dtHistory.Rows.Count > 0 ? true : false;

                        if (check)
                        {
                            Dictionary<string, object> parameterWhere = new Dictionary<string, object>();
                            Dictionary<string, object> parameterTarget = new Dictionary<string, object>();

                            //Set ���� - IP ������Ʈ
                            parameterTarget.Add("PrintValue", cboPort.Text);

                            //WHERE ����
                            parameterWhere.Add("PrintType", cboPrtInfo.Text);

                            int Upresult = _cmm.Update(strTable, parameterWhere, parameterTarget);

                            if (Upresult == -1)
                            {
                                CommonHelper.OpenMessageBox("DB�� ������ �ֽ��ϴ�.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                        }
                        else
                        {
                            Dictionary<int, object> parameter = new Dictionary<int, object>();
                            parameter.Add(0, cboPrtInfo.Text);
                            parameter.Add(1, cboPort.Text);

                            int result = _cmm.Insert(strTable, parameter);

                            if (result == -1)
                            {
                                CommonHelper.OpenMessageBox("������ �����Ͽ����ϴ�.", MessageButtonType.OK, MessageIconType.Error);
                                return;
                            }
                        }
                        //string strLabel = "CT~~CD,~CC^~CT~\t^XA ^PQ1\t^BY2,2.0^FS\t^SEE:UHANGUL.DAT^FS\t^CWQ,E:KFONT3.TTF^CI26^FS\t^LH40,15 ^BQN,3,7 ^FD!!LOT%39Q4-11150%20050002%KR0008%1^FS          ^CF0,40\t^LH240,15 ^AQN,30,30^FD ���^FS             ^CF0,28\t^LH220,60 ^FD Part: ^FS ^LH330,60 ^FD 39Q4-11150^FS\t^LH220,95 ^FD Date: ^FS ^LH330,95 ^FD 2020.05.21^FS\t^LH220,130 ^FD S / N: ^FS ^LH330,130 ^FD 20050002^FS\t^LH220,165 ^FD Vendor: ^FS ^LH330,165 ^AQN,30,30 ^FD �ѱ�����������(��)^FS\t^LH80,200 ^FD LOT%39Q4-11150%20050002%KR0008%1^FS\t^XZ";
                        CommonConnect.SerialName = cboPort.SelectedItem.ToStr();
                        CommonConnect.PrintType = "Serial";
                                                
                        sp.BaudRate = 9600;
                        sp.Parity = Parity.None;
                        sp.DataBits = 8;
                        sp.StopBits = StopBits.One;
                        sp.Encoding = Encoding.Default;
                        sp.PortName = cboPort.Text.ToUpper().Trim();

                        if (CommonConnect.SerialPort == null)
                        {
                            CommonConnect.SerialPort = sp;
                        }
                        else
                        {
                            CommonConnect.SerialPort.Close();
                            CommonConnect.SerialPort.Dispose();
                            CommonConnect.SerialPort = null;

                            CommonConnect.SerialPort = sp;
                        }

                        CommonConnect.SerialPort.Close();

                        CommonConnect.SerialPort.Open();
                        CommonHelper.OpenMessageBox("�ش� ��Ʈ�� ����Ǿ����ϴ�.", MessageButtonType.OK, MessageIconType.Information);
                        //sp.Write(strLabel);
                        dataSendEvent(cboPrtInfo.Text);
                        //sp.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
                //PrintType = CommonConnect.PrintType;
            }
        }

        private void cboPrtInfo_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cboPrtInfo.Text == "TCP/IP")
            {
                layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                layoutControlItem3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else
            {
                layoutControlItem2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                layoutControlItem3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                emptySpaceItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                emptySpaceItem5.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
        }

        private void Setting()
        {
            try
            {
                _cmm = new Common.DataService.DataCommand(_service);

                _service.OpenDB();

                if (_service.IsConn())
                {
                    //���� ������ Ȯ��
                    Dictionary<string, object> parameterForWhere = new Dictionary<string, object>();
                    Dictionary<int, string> parameterForTarget = new Dictionary<int, string>();

                    StringBuilder strQuery = new StringBuilder();
                    strQuery.Append("Select PrintType, PrintValue From PrinttingSetting;");

                    DataSet dsResult = _cmm.Select("", parameterForWhere, parameterForTarget, "M", strQuery.ToStr());

                    DataTable dtHistory = dsResult.Tables[0];

                    int Cnt = dtHistory.Rows.Count;

                    if (dtHistory != null && dtHistory.Rows.Count > 0)
                    {
                        if (Cnt == 2)
                        {
                            for (int i = 0; i < 2; i++)
                            {
                                DataRow row = dtHistory.Rows[i];

                                if (row["PrintType"].ToStr() == "TCP/IP")
                                {
                                    txt_Ip.Text = row["PrintValue"].ToStr();
                                    CommonConnect.PrintIP = txt_Ip.Text;
                                }
                                else
                                {
                                    cboPort.Text = row["PrintValue"].ToStr();
                                }
                            }
                        }
                        else if (Cnt == 1)
                        {
                            if (dtHistory.Rows[0]["PrintType"].Equals("TCP/IP"))
                            {
                                txt_Ip.Text = dtHistory.Rows[0]["PrintValue"].ToStr();
                                CommonConnect.PrintIP = txt_Ip.Text;
                            }
                            else
                            {
                                cboPort.Text = dtHistory.Rows[0]["PrintValue"].ToStr();
                            }
                        }
                    }
                    else
                    {
                        txt_Ip.Text = "";
                        cboPort.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
                CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
            }
            finally
            {
                if (_service.IsConn())
                    _service.CloseDB();
            }
        }
        //Port ����Ʈ
        private void searchPortList()
        {
            string[] portNames;

            portNames = SerialPort.GetPortNames();

            if (portNames.Count() > 0)
            {
                Array.Sort(portNames);

                foreach (var item in portNames)
                {
                    cboPort.Properties.Items.Add(item);
                }
            }

            //cboPort.SelectedIndex = 0;
        }

        //IP ��ȿ���˻�
        public bool IsAvailableIP(string ip)
        {
            // ip ���� null�̸� ���� ��ȯ
            if (ip == null)
                return false;

            // ip ���̰� 15�� �Ѱų� 7���� ������ ���и� ��ȯ    
            if (ip.Length > 15 || ip.Length < 7)
                return false;

            // ���� ����
            int nNumCount = 0;

            // '.' ���� 
            int nDotCount = 0;

            for (int i = 0; i < ip.Length; i++)
            {
                if (ip[i] < '0' || ip[i] > '9')
                {
                    if ('.' == ip[i])
                    {
                        ++nDotCount;
                        nNumCount = 0;
                    }
                    else
                        return false;
                }
                else
                {
                    //'.'�� 4�� �̻��̸� ���� ��ȯ
                    if (++nNumCount > 3)
                        return false;
                }
            }
            // '.' 3�� �ƴϿ��� ���� ��ȯ
            if (nDotCount != 3)
                return false;

            return true;
        }

        private void btn_PopupExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
