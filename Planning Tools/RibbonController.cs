﻿using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;


namespace Planning_Tools
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public void OnButtonPressed_Cells_MyCopy(IRibbonControl control)
        {
            MessageBox.Show("Not yet implemented " + control.Id);
        }
        public void OnButtonPressed_Cells_Paste(IRibbonControl control)
        {
            MessageBox.Show("Not yet implemented " + control.Id);
        }
        public void OnButtonPressed_Cells_Zero(IRibbonControl control, string itemID, int itemIndex)
        {
            MessageBox.Show($"Not yet implemented: control={control} \n idemID={itemID}'\n itemIndex={itemIndex}");
        }
        public void OnButtonPressed_Cells_NullString(IRibbonControl control)
        {
            MessageBox.Show("Not yet implemented ");
        }
        public void OnButtonPressed_Parse_XER(IRibbonControl control)
        {
            var parser = new ParseXER.ParseXER();
            parser.Parse();            
        }
    }
}