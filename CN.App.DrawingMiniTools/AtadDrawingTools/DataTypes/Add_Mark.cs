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
    class Add_Mark
    {
        public Add_Mark(tsd.Bolt dr_bolt)
        {
            tsd.DrawingHandler dh = new DrawingHandler();
            dh.GetDrawingObjectSelector().SelectObject(dr_bolt);
            Tekla.Structures.Model.Operations.Operation.RunMacro("...\\" + "+ADD_MARK");
        }
    }
}
