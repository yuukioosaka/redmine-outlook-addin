using Microsoft.Office.Core;
using System.Windows.Forms;

public class RibbonExtender : IRibbonExtensibility
{
    public string GetCustomUI(string ribbonID)
    {
        return System.IO.File.ReadAllText("Ribbon.xml");
    }

}
