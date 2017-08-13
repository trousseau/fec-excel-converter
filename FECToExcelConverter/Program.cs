using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FECToExcelConverter
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            byte endLoop = 1;
            List<string> lines = new List<String>();

            OpenFileDialog openDialog1 = new OpenFileDialog();
            openDialog1.Filter = "fec files (*.fec)|*.fec";

            SaveFileDialog saveDialog1 = new SaveFileDialog();
            saveDialog1.Filter = "comma separated values (*.csv)|*.csv";

            while (endLoop != 0)
            {
                string path = null;

                FECTools.ParseFecFile(openDialog1,ref path,args, lines);

                endLoop = FECTools.SaveFecFile(lines,path,saveDialog1);
            }
        }
    }
}
