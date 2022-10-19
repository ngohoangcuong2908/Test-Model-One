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
//Khai báo namespace của Tekla
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Drawing;
using Tekla.Structures.Drawing.UI;
using ATADDrawingTools.Functions;
using tsd = Tekla.Structures.Drawing;

namespace ATADDrawingTools.Functions
{
    class ResizeView
    {
        public static tsd.DrawingHandler drawingHandler = new tsd.DrawingHandler(); //tao mot drawinghandler de co the tuong tac voi ban ve.

        public void AutoResizeView(tsd.View view, double offsetFromFrame, double minCutpartDecrease)
        {
            double currentMinCutLength = view.Attributes.Shortening.MinimumLength;
            //Lấy viewbase
            double frameWidth = view.GetDrawing().Layout.SheetSize.Width - 10 - offsetFromFrame;

            //Lấy chiều rộng của viewBase
            //double frameWidth = viewBase.Width - offsetFromFrame;

            for (int i = 0; i < 1000; i++)
            {
                //Lấy chiều rộng của view
                double viewWidth = view.Width;
                //MessageBox.Show(viewWidth.ToString() + "_" + frameWidth.ToString());
                if (viewWidth > frameWidth)
                {
                    view.Select();
                    currentMinCutLength = currentMinCutLength - minCutpartDecrease;
                    view.Attributes.Shortening.MinimumLength = currentMinCutLength;
                    view.Modify();
                }
                else break;
            }
        }
        public void ChangeDrawingMinCutPart(double currentMinCutLength)
        {
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }

            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                               "{" +
                               " public class Script" +
                                   "{" +
                                       " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                       " {" +
               @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
               @"akit.PushButton(""gr_wdraw_on_off"", ""wdraw_dial"");" +
               @"akit.PushButton(""gr_wdraw_view"", ""wdraw_dial"");" +
               @"akit.PushButton(""dv_on_off"", ""wdv_dial"");" +
               @"akit.TabChange(""wdv_dial"", ""contMain"", ""tabAttributes"");" +
               @"akit.ValueChange(""wdv_dial"", ""gr_dv_main_scale"", """ + currentMinCutLength + @""");" +
                @"akit.PushButton(""dv_modify"", ""wdv_dial"");" +
               @" akit.PushButton(""gr_wdraw_ok"", ""wdraw_dial"");" +
               "}}}";
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }
    }
}
