using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel; 

namespace xls2xlsx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        static void Main(string[] args)
        {
            string[] files = args;
            foreach (string filePath in files)
            {
                string ext = Path.GetExtension(filePath).ToLower();
                if (ext == ".xls")
                {
                    runExcel(filePath);
                }
            }
        }

        static void runExcel(string lcFile)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
           // Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(lcFile, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

          
                xlWorkSheet.Activate();
                xlWorkSheet.Application.ActiveWindow.SplitRow = 1;
                xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
       

           
                Excel.Range firstRow = (Excel.Range)xlWorkSheet.Rows[1];
                firstRow.AutoFilter(1,
                    Type.Missing,
                    Excel.XlAutoFilterOperator.xlAnd,
                    Type.Missing,
                    true);
            
                xlWorkSheet.Application.Columns.AutoFit();
            


            try
            {
                xlWorkBook.SaveAs(Filename: lcFile + "x", FileFormat: Excel.XlFileFormat.xlWorkbookDefault);
                deleteOld(lcFile);
            }
            catch (Exception e)
            {
                MessageBox.Show("An error occurred: '{0}'" + e, "File Error", MessageBoxButtons.OK);

            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

           
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        static void deleteOld(string lcFile)
        {
            if (File.Exists(@lcFile + "x"))
            {
                if (File.Exists(@lcFile))
                {
                    File.Delete(@lcFile);
                }

            }
        }

        static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // this.Close();
            Application.Exit();
        }
    }
}
