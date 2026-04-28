using System;
using System.Windows.Forms;

namespace SurveyDataEntry
{
    internal static class Program
    {
        /// <summary>
        /// 해당 애플리케이션의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Windows Forms의 기본 시각적 스타일을 활성화합니다.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 우리가 작성한 Form1(입력기 화면)을 실행하여 화면에 띄웁니다.
            Application.Run(new Form1());
        }
    }
}