using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyAddIn
{
    public partial class Ribbon
    {
        public event Action ButtonClicked;
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (ButtonClicked != null)
            {
                ButtonClicked();
            }
        }
    }
}
