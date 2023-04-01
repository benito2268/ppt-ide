using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerCrisis
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        { 
        }

        public void getText()
        {
            int numSlides = Globals.ThisAddIn.Application.ActivePresentation.Slides.Count;
            string temp = "";
            string filename = "";

            try
            {
                bool firstEl;
                foreach (PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
                {
                    firstEl = true;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                            if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                if(firstEl)
                                {
                                    filename = shape.TextFrame.TextRange.Text;
                                } else
                                {
                                    temp += shape.TextFrame.TextRange.Text;
                                }

                        firstEl = false;
                    }

                    //create a file for this slide
                    create_cpp_file(filename, temp);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("an exception occured " + ex.Message);
            }
        }

        public void create_cpp_file(string filename, string text)
        {
            StreamWriter fout = new StreamWriter("C:\\users\\benst\\pp\\cpp\\" + filename.Trim());
            fout.WriteLine(text);
            fout.Close();
        }


        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
              return new Ribbon1();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
