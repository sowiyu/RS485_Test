using RS485_WinForms_Improved;
namespace RS485_WinForms
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new RS485_ImprovedForm());
        }
    }
}