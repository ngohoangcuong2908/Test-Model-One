using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using tsd = Tekla.Structures.Drawing;

namespace ATADDrawingTools.DataTypes
{
    public class DrawingMark
    {
        public tsd.Drawing Drawing { get; set; }
        public string Prefix { get; set; } // ở đây chính là assmbly position [B001 - 1], prefĩ là B001
        public int SubIndex { get; set; }

    }
}
