using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Linq;

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
                string tempath = null;


                tempath = openFilePath(args, ref path, openDialog1);

                using (TextFieldParser parser1 = new TextFieldParser(path))
                {

                    parser1.SetDelimiters(new string[] { "\x1c" });
                    //parser1.HasFieldsEnclosedInQuotes = true;

                    while (!parser1.EndOfData)
                    {
                        string[] fields = null;

                        //fields = parser1.ReadFields();

                        try
                        {
                            fields = parser1.ReadFields();
                        }
                        catch (MalformedLineException)
                        {

                            break;
                        }


                        if (fields != null)
                        {
                            var newFields = fields
                            //.Select(f => f.Contains(",") ? string.Format("\"{0}\"", f) : f); //only quotes fields with comma
                            .Select(f => f.Replace("\"", "")).Select(f => string.Format("\"{0}\"", f));
                            //.Select(f => string.Format("\"{0}\"", f)); //quotes all fields
                            lines.Add(string.Join(",", newFields));
                        }
                    }
                }

                var formattedCsvToSave = String.Join(Environment.NewLine, lines.Select(x => x));

                string savePath = Path.GetFileNameWithoutExtension(path);



                string saveFilePath = null;

                byte saveCount = 1;

                while (saveCount != 0)
                {
                    saveFilePath = null;
                    saveDialog1.FileName = savePath + "_csv";
                    try
                    {
                        if (saveDialog1.ShowDialog() == DialogResult.Cancel)
                        {
                            Environment.Exit(0);
                        }
                        saveFilePath = saveDialog1.FileName;
                        File.WriteAllText(saveFilePath, formattedCsvToSave);
                        saveCount = 0;
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("File is currently in use");
                    }
                }
                File.WriteAllText(saveFilePath, formattedCsvToSave);

                MessageBox.Show("CSV Saved Successfully");
                endLoop = 0;
            }
        }


        private static string openFilePath(string[] args, ref string path, OpenFileDialog openDialog1)
        {
            string tempath = null;
            byte exitMethod = 1;
            if (args.Length == 0)
            {
                do
                {
                    openDialog1.FileName = String.Empty;
                    if (openDialog1.ShowDialog() == DialogResult.Cancel)
                    {
                        Environment.Exit(0);
                        exitMethod = 0;
                    }
                    tempath = openDialog1.FileName;
                    if (Path.GetExtension(tempath) == ".fec")
                    {
                        path = openDialog1.FileName;
                        exitMethod = 0;
                    }
                    else
                    {
                        MessageBox.Show("Invalid File Extension");
                    }
                }
                while (exitMethod != 0);
            }
            else
            {

                tempath = args[0];

                if (Path.GetExtension(tempath) != ".fec")
                {
                    MessageBox.Show("Invalid File Extension");

                    while (Path.GetExtension(path) != ".fec")
                    {
                        if (openDialog1.ShowDialog() == DialogResult.Cancel)
                        {
                            Environment.Exit(0);
                        }
                        tempath = openDialog1.FileName;
                        if (Path.GetExtension(tempath) == ".fec")
                        {
                            path = openDialog1.FileName;
                            exitMethod = 0;
                        }
                        else
                        {
                            MessageBox.Show("Invalid File Extension");
                        }
                    }
                }
                else
                {
                    path = args[0];
                }
            }

            return tempath;
        }
    }
}
