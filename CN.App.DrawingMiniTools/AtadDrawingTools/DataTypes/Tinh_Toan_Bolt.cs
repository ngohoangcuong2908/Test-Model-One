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
using tsm = Tekla.Structures.Model;
using t3d = Tekla.Structures.Geometry3d;
using tsd = Tekla.Structures.Drawing;
using tsdui = Tekla.Structures.Drawing.UI;

namespace ATADDrawingTools
{
    class Tinh_Toan_Bolt
    {
        public tsm.BoltGroup Bolt_XY { get; set; }//Gồm tất cả điểm của bolt thấy hình tròn
        public tsm.BoltGroup Bolt_X { get; set; }//Gồm tất cả điểm của bolt bolt nằm NGANG (phương X)
        public tsm.BoltGroup Bolt_Y { get; set; }//Gồm tất cả điểm của bolt bolt nằm DỌC (phương Y)
        public tsm.BoltGroup Bolt_skew { get; set; }//Gồm tất cả điểm của bolt bolt nằm XIÊNG
        public t3d.Point PointXmin0 { get; set; } //điểm có tọa độ X và Y nhỏ nhất
        public t3d.Point PointXmin1 { get; set; } //điểm có tọa độ X và Y nhỏ nhất
        public t3d.Point PointXmax0 { get; set; } //điểm có tọa độ X lớn nhất và Y nhỏ nhất
        public t3d.Point PointXmax1 { get; set; } //điểm có tọa độ X lớn nhất và Y nhỏ nhất
        public t3d.Point PointYmin0 { get; set; } //điểm có tọa độ Y nhỏ nhất và X nhỏ nhất
        public t3d.Point PointYmin1 { get; set; } //điểm có tọa độ Y nhỏ nhất và X lớn nhất
        public t3d.Point PointYmax0 { get; set; } //điểm có tọa độ Y lớn nhất và X nhỏ nhất
        public t3d.Point PointYmax1 { get; set; } //điểm có tọa độ Y lớn nhất và X lớn nhất
        public List<t3d.Point> PointListBoltSkew { get; set; }//Gồm tất cả các điểm của bolt Xiêng (không tròn, không nằm ngang, không năm dọc)
        public List<t3d.Point> PointListBolt_XY { get; set; }//Gồm tất cả các điểm của bolt thấy hình tròn
        public List<t3d.Point> PointListBolt_XY_minmaxX { get; set; }//Gồm 2 điểm min và max theo phương X của bolt thấy hình tròn
        public List<t3d.Point> PointListBolt_XY_minmaxY { get; set; }//Gồm 2 điểm min và max theo phương Y của bolt thấy hình tròn
        public List<t3d.Point> PointListBolt_Y { get; set; }//Gồm tất cả điểm của bolt nằm DỌC (phương Y)
        public List<t3d.Point> PointListBolt_Y_minmaxX { get; set; }//Gồm 2 điểm min và max theo phương X của bolt nằm DỌC (phương Y)
        public List<t3d.Point> PointListBolt_Y_minmaxY { get; set; }//Gồm 2 điểm min và max theo phương Y của bolt nằm DỌC (phương Y)
        public List<t3d.Point> PointListBolt_X { get; set; }//Gồm tất cả điểm của bolt nằm DỌC (phương X)
        public List<t3d.Point> PointListBolt_X_minmaxX { get; set; }//Gồm 2 điểm min và max theo phương X của bolt nằm NGANG (phương X)
        public List<t3d.Point> PointListBolt_X_minmaxY { get; set; }//Gồm 2 điểm min và max theo phương Y của bolt nằm NGANG (phương X)
        public Tinh_Toan_Bolt(tsd.View viewcur, tsm.BoltGroup Bolt)
        {
            tsd.ViewBase viewbase = viewcur as tsd.ViewBase;
            //khai báo model hiện hành
            tsm.Model model = new Model();
            //Khai baos workplanehandler
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            //Đưa hệ tọa độ về gốc
            tsm.TransformationPlane OriginalModelPlane = new TransformationPlane();
            wph.SetCurrentTransformationPlane(OriginalModelPlane);
            //Lấy hệ tọa độ của view
            tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
            //xác định hệ tọa độ của bolt
            tsm.TransformationPlane bolt_plane = new TransformationPlane(Bolt.GetCoordinateSystem());
            //Chuyển hệ tọa độ của model về hệ tọa độ phần tử của bolt
            wph.SetCurrentTransformationPlane(bolt_plane);
            //Tính toán điểm để xác định hướng của bolt theo phương Z
            Bolt.Select();
            t3d.Point bolt_sp = Bolt.BoltPositions[0] as t3d.Point;//lấy điểm thứ nhất là tọa độ của Bolt đầu tiên trong boltgroup
            t3d.Point bolt_ep = new t3d.Point(bolt_sp.X, bolt_sp.Y, bolt_sp.Z + 100);//Điểm thứ 2 cùng X, Y chỉ khác Z.
            List<t3d.Point> PointList_XY = new List<t3d.Point>();//Danh sách tọa độ của từng bolt trong bolt group loại XY (bolt tròn thấy được)
            List<t3d.Point> PointList_Y = new List<t3d.Point>();//Danh sách tọa độ của từng bolt trong bolt group loại Y (1 đường nằm dọc)
            List<t3d.Point> PointList_X = new List<t3d.Point>();//Danh sách tọa độ của từng bolt trong bolt group loại Y (1 đường nằm ngang)
            //MessageBox.Show(bolt_sp.ToString() + "/" + bolt_ep.ToString());
            //Chuyển 2 điểm trên về hệ tọa độ global
            t3d.Point g_bolt_sp = bolt_plane.TransformationMatrixToGlobal.Transform(bolt_sp);
            t3d.Point g_bolt_ep = bolt_plane.TransformationMatrixToGlobal.Transform(bolt_ep);
            //Chuyển tập điểm trên về tọa độ của View bản vẽ
            wph.SetCurrentTransformationPlane(viewplane);//Đưa hệ tọa độ làm việc về hệ tọa độ view
            model.CommitChanges();
            t3d.Point dr_bolt_sp = viewplane.TransformationMatrixToLocal.Transform(g_bolt_sp);
            t3d.Point dr_bolt_ep = viewplane.TransformationMatrixToLocal.Transform(g_bolt_ep);
            //Phân loại nhóm bolt
            if (Math.Abs(dr_bolt_sp.X - dr_bolt_ep.X) < 2 && Math.Abs(dr_bolt_sp.Y - dr_bolt_ep.Y) < 2)//Thỏa đk này thì là bolt tròn (bolt nhìn thấy)
            {
                Bolt_XY = Bolt; //Lúc này bolt thỏa điều kiện sẽ đưa vào nhóm Bolt_XY, khai báo ở phía trên public...
                Bolt_XY.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                foreach (t3d.Point p in Bolt_XY.BoltPositions)//Duyệt qua từng bolt trong boltgroup
                {
                    PointList_XY.Add(p);
                }
                if (PointList_XY.Count == 1)
                {
                    PointXmin0 = PointList_XY[0];
                    PointXmax0 = PointList_XY[0];
                    PointYmin0 = PointList_XY[0];
                    PointYmax0 = PointList_XY[0];
                    goto tieptuc;
                }
                List<t3d.Point> PointListOrderBy_X = PointList_XY.OrderBy(point => point.X).ToList(); //để Lấy Xmin0 và Xmin1 (nếu có) 
                t3d.Point pointXmin0 = null; // là điểm X có tọa độ Y nhỏ hơn
                t3d.Point pointXmin1 = null;// là điểm X có tọa độ Y lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                if (Math.Abs(PointListOrderBy_X[0].X - PointListOrderBy_X[1].X) <= 0.01) // nếu mà 2 điểm nhỏ nhất trong danh sách điểm sắp xếp X tăng dần có tọa độ X bằng nhau
                {
                    if (PointListOrderBy_X[0].Y > PointListOrderBy_X[1].Y)// Nếu tọa độ Y của điểm 0 lớn hơn điểm 1 thì điểm 0 là Xmin1, điểm 1 là Xmin0 và ngược lại
                    {
                        pointXmin0 = PointListOrderBy_X[1];
                        pointXmin1 = PointListOrderBy_X[0];
                    }
                    else if (PointListOrderBy_X[0].Y < PointListOrderBy_X[1].Y)
                    {
                        pointXmin0 = PointListOrderBy_X[0];
                        pointXmin1 = PointListOrderBy_X[1];
                    }
                }
                else pointXmin0 = PointListOrderBy_X[0];

                List<t3d.Point> PointListOrderByDescending_X = PointList_XY.OrderByDescending(point => point.X).ToList();//để Lấy Xmax0 và Xmax1 (nếu có) 
                t3d.Point pointXmax0 = null; // là điểm Xmax có tọa độ Y nhỏ hơn
                t3d.Point pointXmax1 = null; // là điểm Xmax có tọa độ Y lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                if (Math.Abs(PointListOrderByDescending_X[0].X - PointListOrderByDescending_X[1].X) <= 0.01) // nếu mà 2 điểm lớn nhất trong danh sách điểm sắp xếp X giảm dần có tọa độ X bằng nhau
                {
                    if (PointListOrderByDescending_X[0].Y > PointListOrderByDescending_X[1].Y)
                    {
                        pointXmax0 = PointListOrderByDescending_X[1];
                        pointXmax1 = PointListOrderByDescending_X[0];
                    }
                    else if (PointListOrderByDescending_X[0].Y < PointListOrderByDescending_X[1].Y)
                    {
                        pointXmax0 = PointListOrderByDescending_X[0];
                        pointXmax1 = PointListOrderByDescending_X[1];
                    }
                }
                else pointXmax0 = PointListOrderByDescending_X[0];

                List<t3d.Point> PointListOrderBy_Y = PointList_XY.OrderBy(point => point.Y).ToList();              //Lấy minY
                t3d.Point pointYmin0 = null; // là điểm Ymin có tọa độ X nhỏ hơn
                t3d.Point pointYmin1 = null;// là điểm Ymin có tọa độ X lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                if (Math.Abs(PointListOrderBy_Y[0].Y - PointListOrderBy_Y[1].Y) <= 0.01) // nếu mà 2 điểm nhỏ nhất trong danh sách điểm sắp xếp Y tăng dần có tọa độ Y bằng nhau
                {
                    if (PointListOrderBy_Y[0].X > PointListOrderBy_Y[1].X)// Nếu tọa độ X của điểm 0 lớn hơn điểm 1 thì điểm 0 là Ymin1, điểm 1 là Ymin0 và ngược lại
                    {
                        pointYmin0 = PointListOrderBy_Y[1];
                        pointYmin1 = PointListOrderBy_Y[0];
                    }
                    else if (PointListOrderBy_Y[0].X < PointListOrderBy_Y[1].X)
                    {
                        pointYmin0 = PointListOrderBy_Y[0];
                        pointYmin1 = PointListOrderBy_Y[1];
                    }
                }
                else pointYmin0 = PointListOrderBy_Y[0];

                List<t3d.Point> PointListOrderByDescending_Y = PointList_XY.OrderByDescending(point => point.Y).ToList();    //Lấy maxY
                t3d.Point pointYmax0 = null; // là điểm Ymin có tọa độ X nhỏ hơn
                t3d.Point pointYmax1 = null;// là điểm Ymin có tọa độ X lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                if (Math.Abs(PointListOrderByDescending_Y[0].Y - PointListOrderByDescending_Y[1].Y) <= 0.01) // nếu mà 2 điểm nhỏ nhất trong danh sách điểm sắp xếp Y tăng dần có tọa độ Y bằng nhau
                {
                    if (PointListOrderByDescending_Y[0].X > PointListOrderByDescending_Y[1].X)// Nếu tọa độ X của điểm 0 lớn hơn điểm 1 thì điểm 0 là Ymin1, điểm 1 là Ymin0 và ngược lại
                    {
                        pointYmax0 = PointListOrderByDescending_Y[1];
                        pointYmax1 = PointListOrderByDescending_Y[0];
                    }
                    else if (PointListOrderByDescending_Y[0].X < PointListOrderByDescending_Y[1].X)
                    {
                        pointYmax0 = PointListOrderByDescending_Y[0];
                        pointYmax1 = PointListOrderByDescending_Y[1];
                    }
                }
                else pointYmax0 = PointListOrderByDescending_Y[0];

                PointXmin0 = pointXmin0;
                if (pointXmin1 != null) PointXmin1 = pointXmin1;
                PointXmax0 = pointXmax0;
                if (pointXmax1 != null) PointXmax1 = pointXmax1;


                PointYmin0 = pointYmin0;
                if (pointYmin1 != null) PointYmin1 = pointYmin1;
                PointYmax0 = pointYmax0;
                if (pointYmax1 != null) PointYmax1 = pointYmax1;
                tieptuc:
                PointListBolt_XY = PointList_XY; //đưa danh sách này ra ngoài PointListBolt_XY
            }
            else if (Math.Abs(dr_bolt_sp.X - dr_bolt_ep.X) < 2 && Math.Abs(dr_bolt_sp.Y - dr_bolt_ep.Y) > 2)//Thỏa đk này thì là bolt nam DOC
            {
                Bolt_Y = Bolt;
                Bolt_Y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                foreach (t3d.Point p in Bolt_Y.BoltPositions)//Duyệt qua từng bolt trong boltgroup
                {
                    PointList_Y.Add(p);
                }
                if (PointList_Y.Count == 1)
                {
                    PointListBolt_Y_minmaxX.Add(PointList_Y[0]); //Gồm 1 điểm min và max theo phương X
                    PointListBolt_Y_minmaxY.Add(PointList_Y[0]);  //Gồm 1 điểm min và max theo phương X
                    goto tieptuc;
                }
                List<t3d.Point> minmaxX = new List<t3d.Point>();
                minmaxX.Add(PointList_Y.OrderBy(point => point.X).ToList()[0]); //Lấy minX
                minmaxX.Add(PointList_Y.OrderByDescending(point => point.X).ToList()[0]); //Lấy maxX
                List<t3d.Point> minmaxY = new List<t3d.Point>();
                minmaxY.Add(PointList_Y.OrderBy(point => point.Y).ToList()[0]); //Lấy minY
                minmaxY.Add(PointList_Y.OrderByDescending(point => point.Y).ToList()[0]); //Lấy maxY
                PointListBolt_Y_minmaxX = minmaxX; //Gồm 2 điểm min và max theo phương X
                PointListBolt_Y_minmaxY = minmaxY; //Gồm 2 điểm min và max theo phương X
            tieptuc:
                PointListBolt_Y = PointList_Y;
            }
            else if (Math.Abs(dr_bolt_sp.X - dr_bolt_ep.X) > 2 && Math.Abs(dr_bolt_sp.Y - dr_bolt_ep.Y) < 2)//Thỏa đk này thì là bolt nam NGANG
            {
                Bolt_X = Bolt;
                Bolt_X.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                foreach (t3d.Point p in Bolt_X.BoltPositions)//Duyệt qua từng bolt trong boltgroup
                {
                    PointList_X.Add(p);
                }
                if (PointList_X.Count == 1)
                {
                    PointListBolt_X_minmaxX.Add(PointList_Y[0]); //Gồm 1 điểm min và max theo phương X
                    PointListBolt_X_minmaxY.Add(PointList_Y[0]);  //Gồm 1 điểm min và max theo phương X
                    goto tieptuc;
                }
                List<t3d.Point> minmaxX = new List<t3d.Point>();
                minmaxX.Add(PointList_X.OrderBy(point => point.X).ToList()[0]); //Lấy minX
                minmaxX.Add(PointList_X.OrderByDescending(point => point.X).ToList()[0]); //Lấy maxX
                List<t3d.Point> minmaxY = new List<t3d.Point>();
                minmaxY.Add(PointList_X.OrderBy(point => point.Y).ToList()[0]); //Lấy minY
                minmaxY.Add(PointList_X.OrderByDescending(point => point.Y).ToList()[0]); //Lấy maxY
                PointListBolt_X_minmaxX = minmaxX; //Gồm 2 điểm min và max theo phương X
                PointListBolt_X_minmaxY = minmaxY; //Gồm 2 điểm min và max theo phương X
            tieptuc:
                PointListBolt_X = PointList_X;
            }
            else
            {
                Bolt_skew = Bolt;
                Bolt_skew.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                foreach (t3d.Point p in Bolt_skew.BoltPositions)//Duyệt qua từng bolt trong boltgroup
                {
                    PointList_X.Add(p);
                }

                List<t3d.Point> minmaxX = new List<t3d.Point>();
                minmaxX.Add(PointList_X.OrderBy(point => point.X).ToList()[0]); //Lấy minX
                minmaxX.Add(PointList_X.OrderByDescending(point => point.X).ToList()[0]); //Lấy maxX
                List<t3d.Point> minmaxY = new List<t3d.Point>();
                minmaxY.Add(PointList_X.OrderBy(point => point.Y).ToList()[0]); //Lấy minY
                minmaxY.Add(PointList_X.OrderByDescending(point => point.Y).ToList()[0]); //Lấy maxY
                PointListBoltSkew = PointList_X; //Gồm 2 điểm min và max theo phương X
                PointListBolt_X_minmaxY = minmaxY; //Gồm 2 điểm min và max theo phương X

            }

        }
    }
}
