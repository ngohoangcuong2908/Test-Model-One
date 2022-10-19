using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//Khai báo namespace của Tekla
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Drawing;
using Tekla.Structures.Drawing.UI;
//Khai báo shortcut cho các Namespace

using tsd = Tekla.Structures.Drawing;



namespace ATADDrawingTools
{
    class Clear_Dimension
    {
        public void ClearDim(tsd.View view)
        {
            tsd.DrawingObjectEnumerator DrObjEnum = view.GetAllObjects(new Type[] { typeof(tsd.StraightDimension) });
            var arrayList = new System.Collections.ArrayList();
            foreach (tsd.DrawingObject DrObj in DrObjEnum)
            {
                DrObj.Select();
                DrObj.Delete();
            }
        }
    }
}
