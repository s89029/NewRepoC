using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Spire.Xls;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using System.Threading;


namespace Jeremy_HDMI_Submit_Helper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string Asa = textBox1.Text + "_" + textBox2.Text;
            DialogResult result;
            result = MessageBox.Show("確定要建立專案資料夾嗎", "", MessageBoxButtons.YesNo);
            FolderBrowserDialog Filepath = new FolderBrowserDialog();
            Filepath.ShowDialog();
            string selectedpath = Filepath.SelectedPath;
            foreach (Control c in groupBox1.Controls)
            {
                if (((System.Windows.Forms.TextBox)c).Text != "")
                {
                    bool dir = false;
                    dir = Directory.Exists(@"" + selectedpath + Asa + @"\" + ((System.Windows.Forms.TextBox)c).Text);
                    if (dir == false)
                    {
                        if (dir == false)
                        {
                            Directory.CreateDirectory(@"" + selectedpath + @"\" + Asa + @"\" + ((System.Windows.Forms.TextBox)c).Text + @"\HDCP");
                            Directory.CreateDirectory(@"" + selectedpath + @"\" + Asa + @"\" + ((System.Windows.Forms.TextBox)c).Text + @"\Impedance");
                            Directory.CreateDirectory(@"" + selectedpath + @"\" + Asa + @"\" + ((System.Windows.Forms.TextBox)c).Text + @"\EVReport");
                            Directory.CreateDirectory(@"" + selectedpath + @"\" + Asa + @"\" + ((System.Windows.Forms.TextBox)c).Text + @"\Protocol");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Fold existed");
                    }
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

            SaveFileDialog savename = new SaveFileDialog();
            savename.DefaultExt = "xlsx";
            savename.Filter = "Excel文件|*.xlsx";
            if (savename.ShowDialog() == DialogResult.OK)
            {
                string fileName = savename.FileName;
                XSSFWorkbook wt = new XSSFWorkbook();
                wt.CreateSheet("name");
                XSSFSheet sheet = (XSSFSheet)wt.GetSheet("name");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("SMName");
                XSSFCellStyle style1 = (XSSFCellStyle)wt.CreateCellStyle();
                XSSFColor mycolor = new XSSFColor(Color.Aqua);
                style1.SetFillBackgroundColor(mycolor);
                sheet.GetRow(0).GetCell(0).CellStyle = style1;
                int j = 0;
                int i = 1;
                foreach (Control c in groupBox1.Controls)
                {
                    if (((TextBox)c).Text != "")
                    {
                        if (j == 0)
                        {
                            foreach (Control d in groupBox2.Controls)
                            {
                                if (((TextBox)d).Text != "")
                                {
                                    sheet.CreateRow(i).CreateCell(j).SetCellValue(((System.Windows.Forms.TextBox)d).Text + "_" + textBox2.Text + "_" + ((System.Windows.Forms.TextBox)c).Text);
                                    i += 1;
                                }
                            }
                            j += 2;
                        }
                        else
                        {
                            i = 1;
                            foreach (Control d in groupBox2.Controls)
                            {
                                if (((TextBox)d).Text != "")
                                {
                                    sheet.GetRow(i).CreateCell(j).SetCellValue(((TextBox)d).Text + "_" + textBox2.Text + "_" + ((TextBox)c).Text);
                                    i += 1;
                                }
                            }
                        }
                    }
                }
                if (savename.FileName != "")
                {
                    try
                    {
                        FileStream file = new FileStream(fileName, FileMode.Create, System.IO.FileAccess.Write);
                        wt.Write(file);
                        file.Close();
                        wt.Close();
                    }
                    catch
                    {
                        MessageBox.Show("存取發生錯誤");
                    }
                }

            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog Filepath = new FolderBrowserDialog();
            Filepath.ShowDialog();
            string selectedpathEV = Filepath.SelectedPath;
            textBox35.Text = selectedpathEV;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog Filepath = new FolderBrowserDialog();
            Filepath.ShowDialog();
            string selectedpathEV = Filepath.SelectedPath;
            textBox36.Text = selectedpathEV;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox35.Text != "" & textBox36.Text != "")
            {
                string FromDirectory25 = @"" + textBox35.Text + @"\25MHz\New Device1";
                string FromDirectory27 = @"" + textBox35.Text + @"\27MHz\New Device1";
                string FromDirectory74 = @"" + textBox35.Text + @"\74MHz\New Device1";
                string FromDirectory148 = @"" + textBox35.Text + @"\148MHz\New Device1";
                string FromDirectory222 = @"" + textBox35.Text + @"\222MHz\New Device1";
                string FromDirectory297 = @"" + textBox35.Text + @"\297MHz\New Device1";
                string FromDirectory371 = @"" + textBox35.Text + @"\371MHz\New Device1";
                string FromDirectory445 = @"" + textBox35.Text + @"\445MHz\New Device1";
                string FromDirectory594 = @"" + textBox35.Text + @"\594MHz\New Device1";
                string ToDirectory = @"" + textBox36.Text;
                string Filter = @"*.html";
                string[] FileList25 = System.IO.Directory.GetFiles(FromDirectory25, Filter);
                string[] FileList27 = System.IO.Directory.GetFiles(FromDirectory27, Filter);
                string[] FileList74 = System.IO.Directory.GetFiles(FromDirectory74, Filter);
                string[] FileList148 = System.IO.Directory.GetFiles(FromDirectory148, Filter);
                string[] FileList222 = System.IO.Directory.GetFiles(FromDirectory222, Filter);
                string[] FileList297 = System.IO.Directory.GetFiles(FromDirectory297, Filter);
                string[] FileList371 = { };
                string[] FileList445 = { };
                try
                {
                    FileList371 = System.IO.Directory.GetFiles(FromDirectory371, Filter);
                }
                catch
                {
                    FileList445 = System.IO.Directory.GetFiles(FromDirectory445, Filter);
                }
                string[] FileList594 = System.IO.Directory.GetFiles(FromDirectory594, Filter);
                foreach (string File in FileList25)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\25MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList27)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\27MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList74)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\74MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList148)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\148MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList222)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\222MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList297)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\297MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList371)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\371MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList445)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\455MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                foreach (string File in FileList594)
                {
                    try
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(File);
                        fi.CopyTo(ToDirectory + @"\597MHz_" + fi.Name);
                    }
                    catch
                    {
                        MessageBox.Show("複製檔案失敗");
                    }
                }
                MessageBox.Show("檔案複製完成");
            }
            else
            {
                MessageBox.Show("來源資料夾與目標資料夾不可為空");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog excelreport = new OpenFileDialog();
            excelreport.Filter = "Excel文件|*.xls";
            excelreport.ShowDialog();
            string excelfilepath = excelreport.FileName;
            textBox37.Text = excelfilepath;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog Filepath = new FolderBrowserDialog();
            Filepath.ShowDialog();
            string pdfpath = Filepath.SelectedPath;
            textBox38.Text = pdfpath;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox37.Text != "" & textBox38.Text != "")
                {
                    foreach (Control c in groupBox1.Controls)
                    {
                        if (((TextBox)c).Text != "")
                        {
                            foreach (Control d in groupBox2.Controls)
                            {
                                if (((TextBox)d).Text != "")
                                {
                                    Workbook excel1 = new Workbook();
                                    excel1.LoadFromFile(textBox37.Text);
                                    Worksheet sheet1 = excel1.Worksheets[0];
                                    Worksheet sheet2 = excel1.Worksheets[1];
                                    sheet1.Range[10, 3].Text = ((TextBox)d).Text + "_" + textBox2.Text + "_" + ((TextBox)c).Text;
                                    sheet1.Range[8, 3].Text = comboBox1.Text ;
                                    sheet2.Range[2, 1].Text = "Project Type: "+ comboBox1.Text + (char)10 + "Product Name: " + ((TextBox)d).Text + "_" + textBox2.Text + "_" + ((TextBox)c).Text + (char)10 + "Family Model:";
                                    excel1.SaveToFile(textBox38.Text + @"\" + ((TextBox)d).Text + "_" + textBox2.Text + "_" + ((TextBox)c).Text + ".xls", FileFormat.Version97to2003);
                                }
                            }
                            
                        }
                        
                    }
                }
                else
                {
                    MessageBox.Show("來源與存檔路徑不可為空");
                }
                MessageBox.Show("Excel產生完成");
            }
            catch
            {
                MessageBox.Show("檔案產生失敗");
            }
        }
        /*public static void ConvertToPdfSdk(string inputFile, string outputFile)
        {
            if (ConvertExtensionToFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);
            
            //Get a ComponentContext
            var xLocalContext =
                Bootstrap.bootstrap();
            //Get MultiServiceFactory
            var xRemoteFactory =
                (XMultiServiceFactory)
                xLocalContext.getServiceManager();
            //Get a CompontLoader
            var aLoader =
                (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
            //Load the sourcefile

            XComponent xComponent = null;
            try
            {
                xComponent = InitDocument(aLoader,
                    PathConverter(inputFile), "_blank");
                //Wait for loading
                while (xComponent == null)
                {
                    Thread.Sleep(1000);
                }

                // save/export the document
                SaveDocument(xComponent, inputFile, PathConverter(outputFile));
            }
            finally
            {
                if (xComponent != null) xComponent.dispose();
            }
        }


        private static XComponent InitDocument(XComponentLoader aLoader, string file, string target)
        {
            var openProps = new PropertyValue[1];
            openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };

            var xComponent = aLoader.loadComponentFromURL(
                file, target, 0,
                openProps);

            return xComponent;
        }

        /// <summary>
        /// 儲存檔案
        /// </summary>
        /// <param name="xComponent">套件</param>
        /// <param name="sourceFile">來源檔案路徑</param>
        /// <param name="destinationFile">目標檔案路徑</param>
        private static void SaveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            var propertyValues = new PropertyValue[2];
            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };
            //// Setting the filter name
            propertyValues[0] = new PropertyValue
            {
                Name = "FilterName",
                Value = new Any(ConvertExtensionToFilterType(Path.GetExtension(sourceFile)))
            };
            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }

        /// <summary>
        /// 檔案路徑字串格式
        /// </summary>
        /// <param name="file">檔案路徑</param>
        /// <returns></returns>
        private static string PathConverter(string file)
        {
            if (string.IsNullOrEmpty(file))
                throw new NullReferenceException("Null or empty path passed to OpenOffice");

            return String.Format("file:///{0}", file.Replace(@"\", "/"));
        }

        /// <summary>
        /// 對應檔案類型
        /// </summary>
        /// <param name="extension">副檔名</param>
        /// <returns></returns>
        public static string ConvertExtensionToFilterType(string extension)
        {
            switch (extension)
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                    return "writer_pdf_Export";
                case ".xls":
                case ".xlsb":
                case ".xlsx":
                case ".ods":
                    return "calc_pdf_Export";
                case ".ppt":
                case ".pptx":
                case ".odp":
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }*/
    }
}
