using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
//Khai báo namespace của Tekla
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Drawing;
using Tekla.Structures.Drawing.UI;

using Tekla.Structures.Solid;
//Khai báo shortcut cho các Namespace
using tsm = Tekla.Structures.Model;
using t3d = Tekla.Structures.Geometry3d;
using tsd = Tekla.Structures.Drawing;
using tsdui = Tekla.Structures.Drawing.UI;
using tss = Tekla.Structures.Solid;

namespace ATADDrawingTools.DataTypes
{
    class CreateDimension
    {
        public static tsd.DrawingHandler myDrawingHandler = new tsd.DrawingHandler(); //tao mot drawinghandler de co the tuong tac voi ban ve.
        public tsdui.Picker c_picker = myDrawingHandler.GetPicker();
        public tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();


        public void CreateStraightDimensionSet_X(tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, new t3d.Vector(0, 1, 0), dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_X_(tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, new t3d.Vector(0, -1, 0), dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_Y(tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, new t3d.Vector(1, 0, 0), dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_Y_(tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, new t3d.Vector(-1, 0, 0), dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_FX(t3d.Point point4, t3d.Point point5, tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            t3d.Vector vx = new Vector(point4.X - point5.X, point4.Y - point5.Y, 0);
            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, vy, dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_FX_(t3d.Point point4, t3d.Point point5, tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            t3d.Vector vx = new Vector(point4.X - point5.X, point4.Y - point5.Y, 0);
            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, -1 * vy, dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_FY(t3d.Point point4, t3d.Point point5, tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            t3d.Vector vx = new Vector(point4.X - point5.X, point4.Y - point5.Y, 0);
            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, vx, dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }
        public void CreateStraightDimensionSet_FY_(t3d.Point point4, t3d.Point point5, tsd.ViewBase viewbase, tsd.PointList pointList, string dimAttributes, int dimSpace)
        {
            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
            t3d.Vector vx = new Vector(point4.X - point5.X, point4.Y - point5.Y, 0);
            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
            tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, -1 * vx, dimSpace, attributes);
            //straightDimensionSet.Select();
            //straightDimensionSet.Distance = dimSpace;
            //straightDimensionSet.Modify();
        }

        public void CreateRadiusDimension(tsd.ViewBase viewbase, t3d.Point point1, t3d.Point point2, t3d.Point point3, double dimSpace, string dimAttributes)
        {
            RadiusDimensionAttributes radiusDimensionAttributes = new RadiusDimensionAttributes(dimAttributes);
            tsd.RadiusDimension radiusDimension = new RadiusDimension(viewbase, point1, point2, point3,dimSpace,radiusDimensionAttributes);
            radiusDimension.Insert();
            //radiusDimension.Select();
            //radiusDimension.Distance = dimSpace;
            //radiusDimension.Modify();
        }
    }
}
