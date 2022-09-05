using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace SAPB1_UI_API
{
    static class Program
    { 

    [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Console.Write("Iniciando");


            SBO_UI_API sBO_UI_API = new SBO_UI_API();

            Application.Run();
            //Application.Run(new Form1());
        }
    }
}
