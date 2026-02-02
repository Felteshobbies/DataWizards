using System;
using System.Windows.Forms;

namespace DataWizard.UI
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Prüfe auf Command-Line Arguments für Watcher-Modus
            if (args.Length > 0 && args[0] == "--watch")
            {
                // Starte direkt im Watcher-Modus
                fMain form = new fMain();
                form.Load += (s, e) =>
                {
                    if (args.Length >= 2)
                    {
                        // TODO: Setze Watch-Ordner aus args[1]
                        // form.SetWatchFolder(args[1]);
                    }
                };
                Application.Run(form);
            }
            else
            {
                Application.Run(new fMain());
            }
        }
    }
}
