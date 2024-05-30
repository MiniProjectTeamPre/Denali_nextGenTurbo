using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Denali_nextGenTurbo {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new Form1());
            return;

            Application.Run(new TeamPrecision.PRISM.frmUserLogin());

            if (TeamPrecision.PRISM.cSettingValues.EmployeeID != "") {
                
                Application.Run(new Form1());
            }
        }

        static void UnhandledException(object sender, UnhandledExceptionEventArgs e) {
            var exc = e.ExceptionObject as Exception;
            if (exc != null) {
                File.WriteAllText("ErrorProgram.txt", exc.ToString());
            }
            Environment.Exit(-1);
        }
    }
}
