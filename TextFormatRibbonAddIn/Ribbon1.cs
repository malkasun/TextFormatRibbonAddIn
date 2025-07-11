using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace TextFormatRibbonAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //private void btnStrikethrough_Click(object sender, RibbonControlEventArgs e)
        //{
        //    var app = Globals.ThisAddIn.Application;
        //    PowerPoint.Selection selection = app.ActiveWindow.Selection;

        //    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
        //    {
        //        PowerPoint.TextRange textRange = selection.TextRange;
        //        textRange.Font.Strikethrough = Office.MsoTriState.msoTrue;
        //    }

        //}

        //private void btnDoubleStrikethrough_Click(object sender, RibbonControlEventArgs e)
        //{
        //    PowerPoint.Application app = Globals.ThisAddIn.Application;
        //    PowerPoint.Selection selection = app.ActiveWindow.Selection;

        //    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
        //    {
        //        PowerPoint.TextRange textRange = selection.TextRange;
        //        textRange.Font.DoubleStrikethrough = Microsoft.Office.Core.MsoTriState.msoTrue;
        //    }

        //}

        private void btnSuperscript_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange textRange = selection.TextRange;
                textRange.Font.Superscript = Office.MsoTriState.msoTrue;
            }

        }

        private void btnSubscript_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            PowerPoint.Selection selection = app.ActiveWindow.Selection;

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange textRange = selection.TextRange;
                textRange.Font.Subscript = Office.MsoTriState.msoTrue;
            }

        }

        //private void btnSmallCaps_Click(object sender, RibbonControlEventArgs e)
        //{
        //    var app = Globals.ThisAddIn.Application;
        //    PowerPoint.Selection selection = app.ActiveWindow.Selection;

        //    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
        //    {
        //        PowerPoint.TextRange textRange = selection.TextRange;
        //        textRange.Font.SmallCaps = Office.MsoTriState.msoTrue;
        //    }

        //}

        //private void btnAllCaps_Click(object sender, RibbonControlEventArgs e)
        //{
        //    var app = Globals.ThisAddIn.Application;
        //    PowerPoint.Selection selection = app.ActiveWindow.Selection;

        //    if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
        //    {
        //        PowerPoint.TextRange textRange = selection.TextRange;
        //        textRange.Font.AllCaps = Office.MsoTriState.msoTrue;
        //    }
        //}

        private void btnEqualizeHeight_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Double strikethrough is not supported in PowerPoint.");
        }
    }
}
