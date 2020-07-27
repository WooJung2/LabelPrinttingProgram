#region ================= Common Util =================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Data;
using DevExpress.XtraSplashScreen;
using System.Data.OleDb;
using System.Globalization;
using System.Text.RegularExpressions;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using Common;

namespace Common
{
    /// <summary>
    /// ���� Util
    /// </summary>
    public static class CommonHelper
    {
        /// <summary>
        /// Log ����
        /// </summary>
        /// <param name="eventId"></param>
        /// <param name="data"></param>
        private static void SetLog(string eventId, string data, bool useMessageBox = true)
        {
            try
            {
                DirectoryInfo di;
                string sLog = string.Empty;

                sLog += "==========================================================================" + Environment.NewLine;
                sLog += "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "]" + Environment.NewLine;
                sLog += "[" + eventId + "]" + Environment.NewLine;
                sLog += data + Environment.NewLine;
                sLog += "==========================================================================" + Environment.NewLine;

                //Log ���� �ִ��� Ȯ��
                di = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + @"\Log");
                //������ ����
                if (di.Exists == false) di.Create();

                //Log �ۼ�
                File.AppendAllText(System.Windows.Forms.Application.StartupPath + @"\Log\" + DateTime.Now.ToString("yyyyMMdd") + ".log", sLog);


                //Message Box ���
                if (useMessageBox)
                    OpenMessageBox(data, MessageButtonType.OK, MessageIconType.Error);
            }
            catch { }
        }

        /// <summary>
        /// Message Box
        /// </summary>
        /// <param name="Message">��� �޽���</param>
        /// <param name="boxButtonType">��ư Ÿ��</param>
        /// <param name="icon">������ Ÿ��</param>
        /// <returns></returns>
        public static DialogResult OpenMessageBox(string Message, MessageButtonType boxButtonType = MessageButtonType.OK, MessageIconType icon = MessageIconType.Information)
        {

            if (Class.msgbox == null)
                Class.msgbox = new Common.MessageForm();

            return Class.msgbox.OpenMessageForm(Message, boxButtonType, icon);
        }
        /// <summary>
        /// ���� ó�� �� �޽�¡ ó��
        /// </summary>
        /// <param name="FormName"></param>
        /// <param name="ErrorMessage"></param>
        public static void SetLogAndMessage(string FormName, string ErrorMessage, bool UseMessageBox = true)
        {
            CommonHelper.SetLog(FormName, ErrorMessage);
        }
    }
}

#endregion