using System;
using System.Windows.Forms;

namespace UartSwitchControl
{
    internal static class Program
    {
        /// <summary>
        /// ���ε{�����D�n�i�J�I
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());  // �Ұʧڭ̪����
        }
    }
}
