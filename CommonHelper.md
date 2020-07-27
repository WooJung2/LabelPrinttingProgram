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
    /// 공통 Util
    /// </summary>
    public static class CommonHelper
    {
        /// <summary>
        /// Log 생성
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

                //Log 폴더 있는지 확인
                di = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + @"\Log");
                //없으면 생성
                if (di.Exists == false) di.Create();

                //Log 작성
                File.AppendAllText(System.Windows.Forms.Application.StartupPath + @"\Log\" + DateTime.Now.ToString("yyyyMMdd") + ".log", sLog);


                //Message Box 출력
                if (useMessageBox)
                    OpenMessageBox(data, MessageButtonType.OK, MessageIconType.Error);
            }
            catch { }
        }

        /// <summary>
        /// Message Box
        /// </summary>
        /// <param name="Message">출력 메시지</param>
        /// <param name="boxButtonType">버튼 타입</param>
        /// <param name="icon">아이콘 타입</param>
        /// <returns></returns>
        public static DialogResult OpenMessageBox(string Message, MessageButtonType boxButtonType = MessageButtonType.OK, MessageIconType icon = MessageIconType.Information)
        {

            if (Class.msgbox == null)
                Class.msgbox = new Common.MessageForm();

            return Class.msgbox.OpenMessageForm(Message, boxButtonType, icon);
        }
        /// <summary>
        /// 오류 처리 및 메시징 처리
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