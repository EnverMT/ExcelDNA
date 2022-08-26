using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;


namespace Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public void OnButtonPressed_Cells_MyCopy(IRibbonControl control)
        {   
            DataWriter.WriteData();
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
            Planning_Tools.ParseXER.Parse();            
        }
    }
}