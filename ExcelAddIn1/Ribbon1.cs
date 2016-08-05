using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Crabyter
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveCell.NumberFormatLocal = "@";
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            int number = Globals.ThisAddIn.Application.ActiveCell.Column;
            Globals.ThisAddIn.Application.Cells[1][1] = number;
        }
    }
}
