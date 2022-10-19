using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Drawing.UI;
using Tekla.Structures.Dialog; // để ApplicationFormBase không bị lỗi
using Tekla.Structures.Solid;
using tsmui = Tekla.Structures.Model.UI;
using t3d = Tekla.Structures.Geometry3d;
using tsd = Tekla.Structures.Drawing;
using tsm = Tekla.Structures.Model;
using tsdui = Tekla.Structures.Drawing.UI;
using tss = Tekla.Structures.Solid;

using ATADDrawingTools.DataTypes;
using System.Text.RegularExpressions;
using System.Collections;
using System.IO;
using Newtonsoft.Json;

namespace ATADDrawingTools.Functions
{
    class GroupNSFS
    {
        public tsm.Part PartLeft { get; }
        public tsm.Part PartRight { get; }
        public string PartMarkLeft { get; }
        public string PartMarkRight { get; }
        public GroupNSFS(tsm.Part Left, tsm.Part Right)
        {
            PartLeft = Left;
            PartRight = Right;
            Function function = new Function();
            string partMarkLeft = function.GetPartPos(Left);
            string partMarkRight = function.GetPartPos(Right);
            PartMarkLeft = partMarkLeft;
            PartMarkRight = partMarkRight;
        }
    }
}
