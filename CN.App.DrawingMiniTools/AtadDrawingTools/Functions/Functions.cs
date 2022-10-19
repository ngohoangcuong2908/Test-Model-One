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
using ATADDrawingTools.Functions;

namespace ATADDrawingTools.Functions
{
    public class Function
    {
        public void SwapDirection(ref t3d.Vector vector, DirectionEnum directionEnum)
        {
            t3d.Point p = new t3d.Point();
            vector.Normalize(100);
            p.Translate(vector.X, vector.Y, vector.Z);
            switch (directionEnum)
            {
                case DirectionEnum.Top:
                    if (p.Y < 0) vector = -1 * vector;
                    break;
                case DirectionEnum.Bottom:
                    if (p.Y > 0) vector = -1 * vector;
                    break;
                case DirectionEnum.Left:
                    if (p.X > 0) vector = -1 * vector;
                    break;
                case DirectionEnum.Right:
                    if (p.X < 0) vector = -1 * vector;
                    break;
            }
        }
        public double GetWidth(tsm.Part part)
        {
            double width = 0;
            part.GetReportProperty("PROFILE.WIDTH", ref width);
            return width;
        }
        public double GetHeight(tsm.Part part)
        {
            double height = 0;
            part.GetReportProperty("PROFILE.HEIGHT", ref height);
            return height;
        }
        public double GetLength(tsm.Part part)
        {
            double length = 0;
            part.GetReportProperty("LENGTH", ref length);
            return length;
        }
        public string GetProfileType(tsm.Part part)
        {
            string partProfileType = string.Empty;
            part.GetReportProperty("PROFILE_TYPE", ref partProfileType);
            return partProfileType;
        }

        public string GetPartPos(tsm.Part part)
        {
            string partPos = string.Empty;
            part.GetReportProperty("PART_POS", ref partPos);
            return partPos;
        }

        public void InsertSymbol(tsd.ViewBase viewBase, t3d.Point p)
        {
            tsd.Symbol symbol = new Symbol(viewBase, p, new tsd.SymbolInfo("xsteel", 1));
            symbol.Insert();
        }
        public void InsertLine(tsd.ViewBase viewBase, t3d.Point p1, t3d.Point p2)
        {
            tsd.Line line = new tsd.Line(viewBase, p1, p2);
            line.Insert();
        }

        public t3d.Point FindIntersection(tsd.Line lineA, tsd.Line lineB, double tolerance = 0.001)
        {
            double x1 = lineA.StartPoint.X, y1 = lineA.StartPoint.Y;
            double x2 = lineA.EndPoint.X, y2 = lineA.EndPoint.Y;

            double x3 = lineB.StartPoint.X, y3 = lineB.StartPoint.Y;
            double x4 = lineB.EndPoint.X, y4 = lineB.EndPoint.Y;

            // equations of the form x = c (two vertical lines)
            if (Math.Abs(x1 - x2) < tolerance && Math.Abs(x3 - x4) < tolerance && Math.Abs(x1 - x3) < tolerance)
            {
                throw new Exception("Both lines overlap vertically, ambiguous intersection points.");
            }

            //equations of the form y=c (two horizontal lines)
            if (Math.Abs(y1 - y2) < tolerance && Math.Abs(y3 - y4) < tolerance && Math.Abs(y1 - y3) < tolerance)
            {
                throw new Exception("Both lines overlap horizontally, ambiguous intersection points.");
            }

            //equations of the form x=c (two vertical lines)
            if (Math.Abs(x1 - x2) < tolerance && Math.Abs(x3 - x4) < tolerance)
            {
                return default(t3d.Point);
            }

            //equations of the form y=c (two horizontal lines)
            if (Math.Abs(y1 - y2) < tolerance && Math.Abs(y3 - y4) < tolerance)
            {
                return default(t3d.Point);
            }

            //general equation of line is y = mx + c where m is the slope
            //assume equation of line 1 as y1 = m1x1 + c1 
            //=> -m1x1 + y1 = c1 ----(1)
            //assume equation of line 2 as y2 = m2x2 + c2
            //=> -m2x2 + y2 = c2 -----(2)
            //if line 1 and 2 intersect then x1=x2=x & y1=y2=y where (x,y) is the intersection point
            //so we will get below two equations 
            //-m1x + y = c1 --------(3)
            //-m2x + y = c2 --------(4)

            double x, y;

            //lineA is vertical x1 = x2
            //slope will be infinity
            //so lets derive another solution
            if (Math.Abs(x1 - x2) < tolerance)
            {
                //compute slope of line 2 (m2) and c2
                double m2 = (y4 - y3) / (x4 - x3);
                double c2 = -m2 * x3 + y3;

                //equation of vertical line is x = c
                //if line 1 and 2 intersect then x1=c1=x
                //subsitute x=x1 in (4) => -m2x1 + y = c2
                // => y = c2 + m2x1 
                x = x1;
                y = c2 + m2 * x1;
            }
            //lineB is vertical x3 = x4
            //slope will be infinity
            //so lets derive another solution
            else if (Math.Abs(x3 - x4) < tolerance)
            {
                //compute slope of line 1 (m1) and c2
                double m1 = (y2 - y1) / (x2 - x1);
                double c1 = -m1 * x1 + y1;

                //equation of vertical line is x = c
                //if line 1 and 2 intersect then x3=c3=x
                //subsitute x=x3 in (3) => -m1x3 + y = c1
                // => y = c1 + m1x3 
                x = x3;
                y = c1 + m1 * x3;
            }
            //lineA & lineB are not vertical 
            //(could be horizontal we can handle it with slope = 0)
            else
            {
                //compute slope of line 1 (m1) and c2
                double m1 = (y2 - y1) / (x2 - x1);
                double c1 = -m1 * x1 + y1;

                //compute slope of line 2 (m2) and c2
                double m2 = (y4 - y3) / (x4 - x3);
                double c2 = -m2 * x3 + y3;

                //solving equations (3) & (4) => x = (c1-c2)/(m2-m1)
                //plugging x value in equation (4) => y = c2 + m2 * x
                x = (c1 - c2) / (m2 - m1);
                y = c2 + m2 * x;

                //verify by plugging intersection point (x, y)
                //in orginal equations (1) & (2) to see if they intersect
                //otherwise x,y values will not be finite and will fail this check
                if (!(Math.Abs(-m1 * x + y - c1) < tolerance
                    && Math.Abs(-m2 * x + y - c2) < tolerance))
                {
                    return default(t3d.Point);
                }
            }

            //x,y can intersect outside the line segment since line is infinitely long
            //so finally check if x, y is within both the line segments
            if (IsInsideLine(lineA, x, y) &&
                IsInsideLine(lineB, x, y))
            {
                return new t3d.Point { X = x, Y = y };
            }

            //return default null (no intersection)
            return default(t3d.Point);

        }

        // Returns true if given point(x,y) is inside the given line segment
        private static bool IsInsideLine(tsd.Line line, double x, double y)
        {
            return (x >= line.StartPoint.X && x <= line.EndPoint.X
                        || x >= line.EndPoint.X && x <= line.StartPoint.X)
                   && (y >= line.StartPoint.Y && y <= line.EndPoint.Y
                        || y >= line.EndPoint.Y && y <= line.StartPoint.Y);
        }

    }
}
