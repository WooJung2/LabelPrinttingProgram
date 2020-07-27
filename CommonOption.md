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
        public const string FontFailmyName = "나눔고딕";
        public const float FontEmSize = 9f;
        public const System.Drawing.FontStyle FontStyle = System.Drawing.FontStyle.Regular;

        public const string POSFontFailmyName = "나눔고딕";
        public const float POSFontEmSize = 10f;
        public const System.Drawing.FontStyle POSFontStyle = System.Drawing.FontStyle.Bold;        
    }

    public enum MessageIconType
    {
        //
        // 요약:
        //     메시지 상자에는 빨간색 배경의 원 안에 흰색 X가 포함된 기호가 있습니다.
        Error = 16,
        //
        // 요약:
        //     메시지 상자에는 원 안에 물음표가 포함된 기호가 있습니다.물음표 메시지 아이콘은 특정 메시지 유형을 명확하게 나타내지 않고 메시지를
        //     질문으로 구문 분석하면 모든 메시지 유형에 적용되므로 더 이상 권장되지 않습니다.또한 사용자가 메시지 기호 물음표를 도움말 정보와
        //     혼동할 수 있으므로메시지 상자에 이 물음표 메시지 기호를 사용하지 마십시오.시스템에서는 이전 버전과의 호환성을 위해 이 물음표 메시지
        //     기호를 계속 지원합니다.
        Question = 32,
        //
        // 요약:
        //     메시지 상자에는 노란색 배경의 삼각형 안에 느낌표가 포함된 기호가 있습니다.
        Warning = 48,
        //
        // 요약:
        //     메시지 상자에는 원 안에 소문자 i가 포함된 기호가 있습니다.
        Information = 64,
        //
        // 요약:
        //     메시지 상자에는 원 안에 소문자 i가 포함된 기호가 있습니다.
        Asterisk = 64,
    }

    public enum MessageButtonType
    {
        // 요약:
        //     메시지 상자에 확인 단추가 있습니다.
        OK = 0,
        //
        // 요약:
        //     메시지 상자에 확인 및 취소 단추가 있습니다.
        OKCancel = 1,
        //
        // 요약:
        //     메시지 상자에 예, 아니요 및 취소 단추가 있습니다.
        YesNoCancel = 3,
        //
        // 요약:
        //     메시지 상자에 예 및 아니요 단추가 있습니다.
        YesNo = 4,
        //
        // 요약:
        //     메시지 상자에 다시 시도 및 취소 단추가 있습니다.
        RetryCancel = 5,
    }

    public enum EditType
    {
        /// <summary>
        /// 문자열
        /// </summary>
        STRING = 0,

        /// <summary>
        /// Integer
        /// </summary>
        INT = 1,
        /// <summary>
        /// Double 소수점 없음
        /// </summary>
        DOUBLE_n0 = 2,
        /// <summary>
        /// Double 소수점2자리
        /// </summary>
        DOUBLE_n2 = 3,

        /// <summary>
        /// CUSTOM 1
        /// </summary>
        LOT_ID = 99,
    }

}
