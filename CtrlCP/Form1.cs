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

using Microsoft.Office.Core;
using Excel =Microsoft.Office.Interop.Excel;
namespace CtrlCP
{
    public partial class Form1 : Form
    {
        string sourcepath="";
        string targetlocation = "";
        string filename = "";
        public Form1()
        {
            InitializeComponent();
            txtSource.ReadOnly = true;
            txtTarget.ReadOnly = true;

        }

        private void btnSource_Click(object sender, EventArgs e)
        {
            DialogResult result = FileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                //Source File Name
                filename = FileDialog.FileName.Split('\\').Last();
                //Source full path
                sourcepath = FileDialog.FileName;
                //Displaying the source path in the text box
                txtSource.Paste(sourcepath);
            }
        }

        private void btnTarget_Click(object sender, EventArgs e)
        {
            DialogResult result = FolderBrowse.ShowDialog(); 
            if (result == DialogResult.OK) 
            {
                //Destination path
                targetlocation = FolderBrowse.SelectedPath;
                //Displaying the source path in the text box
                txtTarget.Paste(targetlocation);

            }

        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            FileInfo srcfileinfo = new FileInfo(sourcepath);
            srcfileinfo.CopyTo(targetlocation+"\\"+filename,false);
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "Source File Name";
            oSheet.Cells[1, 2] = "Source Path";
            oSheet.Cells[1, 3] = "Destination Path";

            //Format A1:C1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "C1").Font.Bold = true;
            oSheet.get_Range("A1", "C1").VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Create an array to multiple values at once.
            string[] values = new string[] { filename, sourcepath, targetlocation };
            //Fill A2:C2 with an array of values
            oSheet.get_Range("A2", "C2").Value2 = values;


            //AutoFit columns A:C.
            oRng = oSheet.get_Range("A1", "C1");
            oRng.EntireColumn.AutoFit();

            oXL.Visible = false;
            oXL.UserControl = false;
            oWB.SaveAs("d:\\CopyPaste.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
            MessageBox.Show("Excel file created , you can find the file d:\\CopyPaste.xls");




        }

    }
}
