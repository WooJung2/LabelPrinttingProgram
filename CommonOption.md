using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
namespace Common
{
    public static class DefaultOption
    {
        public const string FontFailmyName = "�������";
        public const float FontEmSize = 9f;
        public const System.Drawing.FontStyle FontStyle = System.Drawing.FontStyle.Regular;

        public const string POSFontFailmyName = "�������";
        public const float POSFontEmSize = 10f;
        public const System.Drawing.FontStyle POSFontStyle = System.Drawing.FontStyle.Bold;        
    }

    public enum MessageIconType
    {
        //
        // ���:
        //     �޽��� ���ڿ��� ������ ����� �� �ȿ� ��� X�� ���Ե� ��ȣ�� �ֽ��ϴ�.
        Error = 16,
        //
        // ���:
        //     �޽��� ���ڿ��� �� �ȿ� ����ǥ�� ���Ե� ��ȣ�� �ֽ��ϴ�.����ǥ �޽��� �������� Ư�� �޽��� ������ ��Ȯ�ϰ� ��Ÿ���� �ʰ� �޽�����
        //     �������� ���� �м��ϸ� ��� �޽��� ������ ����ǹǷ� �� �̻� ������� �ʽ��ϴ�.���� ����ڰ� �޽��� ��ȣ ����ǥ�� ���� ������
        //     ȥ���� �� �����Ƿθ޽��� ���ڿ� �� ����ǥ �޽��� ��ȣ�� ������� ���ʽÿ�.�ý��ۿ����� ���� �������� ȣȯ���� ���� �� ����ǥ �޽���
        //     ��ȣ�� ��� �����մϴ�.
        Question = 32,
        //
        // ���:
        //     �޽��� ���ڿ��� ����� ����� �ﰢ�� �ȿ� ����ǥ�� ���Ե� ��ȣ�� �ֽ��ϴ�.
        Warning = 48,
        //
        // ���:
        //     �޽��� ���ڿ��� �� �ȿ� �ҹ��� i�� ���Ե� ��ȣ�� �ֽ��ϴ�.
        Information = 64,
        //
        // ���:
        //     �޽��� ���ڿ��� �� �ȿ� �ҹ��� i�� ���Ե� ��ȣ�� �ֽ��ϴ�.
        Asterisk = 64,
    }

    public enum MessageButtonType
    {
        // ���:
        //     �޽��� ���ڿ� Ȯ�� ���߰� �ֽ��ϴ�.
        OK = 0,
        //
        // ���:
        //     �޽��� ���ڿ� Ȯ�� �� ��� ���߰� �ֽ��ϴ�.
        OKCancel = 1,
        //
        // ���:
        //     �޽��� ���ڿ� ��, �ƴϿ� �� ��� ���߰� �ֽ��ϴ�.
        YesNoCancel = 3,
        //
        // ���:
        //     �޽��� ���ڿ� �� �� �ƴϿ� ���߰� �ֽ��ϴ�.
        YesNo = 4,
        //
        // ���:
        //     �޽��� ���ڿ� �ٽ� �õ� �� ��� ���߰� �ֽ��ϴ�.
        RetryCancel = 5,
    }

    public enum EditType
    {
        /// <summary>
        /// ���ڿ�
        /// </summary>
        STRING = 0,

        /// <summary>
        /// Integer
        /// </summary>
        INT = 1,
        /// <summary>
        /// Double �Ҽ��� ����
        /// </summary>
        DOUBLE_n0 = 2,
        /// <summary>
        /// Double �Ҽ���2�ڸ�
        /// </summary>
        DOUBLE_n2 = 3,

        /// <summary>
        /// CUSTOM 1
        /// </summary>
        LOT_ID = 99,
    }

}
