using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATADDrawingTools.DataTypes
{
    public enum DirectionEnum // phương của dim
    {
        Top, Bottom, Left, Right
    }
    public enum EndPlateEnum //end plate phía Trái (start), hoặc Phải (End)
    {
        Start, End
    }
    public enum PartPositionEnum // end plate nằm ngang, dọc, nghiêng trái, nghiêng phải
    {
        Horizontal, Vertical, Left, Right
    }
}
