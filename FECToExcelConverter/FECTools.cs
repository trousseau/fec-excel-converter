using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Linq;

namespace FECToExcelConverter
{
    class FECTools
    {
        public static string OpenFilePath(string[] args, ref string path, OpenFileDialog openDialog1)
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

                try
                {
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
                catch (Exception e)
                {

                    MessageBox.Show(e.Message);
                }
            }
            return tempath;
        }

        public static string ParseFecFile(OpenFileDialog openDialog1, ref string path, string[] args, List<string> lines)
        {
            string tempath = OpenFilePath(args, ref path, openDialog1);

            try
            {
                using (TextFieldParser parser1 = new TextFieldParser(path))
                {
                    parser1.SetDelimiters(new string[] { "\x1c" });

                    while (!parser1.EndOfData)
                    {
                        string[] fields = null;

                        try
                        {
                            fields = parser1.ReadFields();
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(e.Message);
                            break;
                        }

                        if (fields != null)
                        {
                            var newFields = fields
                            .Select(f => f.Replace("\"", "")).Select(f => string.Format("\"{0}\"", f));
                            lines.Add(string.Join(",", newFields));
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return tempath;
        }

        public static byte SaveFecFile(List<string> lines,string path, SaveFileDialog saveDialog1)
        {
            var formattedCsvToSave = string.Empty;
            string savePath = string.Empty;
            string saveFilePath = null;
            byte saveCount = 1;

            try
            {
                formattedCsvToSave = String.Join(Environment.NewLine, lines.Select(x => x));
                savePath = Path.GetFileNameWithoutExtension(path);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

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
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }

            try
            {
                File.WriteAllText(saveFilePath, formattedCsvToSave);
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message);
            }

            MessageBox.Show("CSV Saved Successfully");
            return 0;
        }
    }
}
