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
//Khai báo shortcut cho các Namespace
using ATADDrawingTools.Functions;
using tsd = Tekla.Structures.Drawing;

namespace ATADDrawingTools.Functions

{
    public class ClearDrawingObjects
    {
        public DrawingHandler drawingHandler = new DrawingHandler();
        public void ClearDim(tsd.View view)
        {
            try
            {
                tsd.DrawingObjectEnumerator DrObjEnum = view.GetAllObjects(new Type[] { typeof(tsd.DimensionBase) });
                var arrayList = new System.Collections.ArrayList();
                foreach (tsd.DrawingObject DrObj in DrObjEnum)
                    arrayList.Add(DrObj);
                if (arrayList.Count != 0)
                {
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(arrayList, false);
                    Delete();
                }
            }
            catch
            {
                MessageBox.Show("Cannot Clear Dimension. Try to update drawing before run tool");
            }

        }

        public void ClearPartMark(tsd.View view)
        {
            try
            {
                tsd.DrawingObjectEnumerator DrObjEnum = view.GetAllObjects(new Type[] { typeof(tsd.MarkBase) });
                var arrayList = new System.Collections.ArrayList();
                foreach (tsd.DrawingObject DrObj in DrObjEnum)
                    arrayList.Add(DrObj);
                if (arrayList.Count != 0)
                {
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(arrayList, false);
                    Delete();
                }
            }
            catch
            {
                MessageBox.Show("Cannot Clear Mark. Try to update drawing before run tool");
            }

        }

        public void ClearText(tsd.View view)
        {
            tsd.DrawingObjectEnumerator DrObjEnum = view.GetAllObjects(new Type[] { typeof(tsd.Text) });
            var arrayList = new System.Collections.ArrayList();
            foreach (tsd.DrawingObject DrObj in DrObjEnum)
            {
                arrayList.Add(DrObj);
            }
            drawingHandler.GetDrawingObjectSelector().SelectObjects(arrayList, false);
            Delete();
        }
        public void ClearCloud(tsd.View view)
        {
            tsd.DrawingObjectEnumerator DrObjEnum = view.GetAllObjects(new Type[] { typeof(tsd.Cloud) });
            var arrayList = new System.Collections.ArrayList();
            foreach (tsd.DrawingObject DrObj in DrObjEnum)
            {
                arrayList.Add(DrObj);
            }
            drawingHandler.GetDrawingObjectSelector().SelectObjects(arrayList, false);
            Delete();
        }
        public void ClearRectangle(tsd.View view)
        {
            tsd.DrawingObjectEnumerator DrObjEnum = view.GetAllObjects(new Type[] { typeof(tsd.Rectangle) });
            var arrayList = new System.Collections.ArrayList();
            foreach (tsd.DrawingObject DrObj in DrObjEnum)
            {
                arrayList.Add(DrObj);
            }
            drawingHandler.GetDrawingObjectSelector().SelectObjects(arrayList, false);
            Delete();
        }

        public void Delete()
        {
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
            //Create a script as string
            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                            "{" +
                            " public class Script" +
                                "{" +
                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                    " {" +
            @"akit.Callback(""acmd_delete_selected_dr"", """", ""main_frame"");" +
            "}}}";
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }
    }
}
