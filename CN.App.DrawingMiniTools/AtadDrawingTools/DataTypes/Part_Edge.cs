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

namespace ATADDrawingTools
{
    class Part_Edge
    {
        public static tsd.DrawingHandler dh = new DrawingHandler();//Khai báo kết nối với bản vẽ qua Drawinghandler
        public static tsm.Model myModel = new tsm.Model(); //Ket noi voi model hien hanh

        //Lấy thuộc tính của Part này ra ngoài để dùng
        public t3d.Point PointXmin0 = null;// là điểm có tọa độ Y nhỏ hơn
        public t3d.Point PointXmin1 = null;// là điểm có tọa độ Y lớn hơn
        public t3d.Point PointXmax0 = null;// là điểm có tọa độ Y nhỏ hơn
        public t3d.Point PointXmax1 = null;// là điểm có tọa độ Y lớn hơn
        public t3d.Point PointYmin0 = null;// là điểm có tọa độ X nhỏ hơn
        public t3d.Point PointYmin1 = null;// là điểm có tọa độ X lớn hơn
        public t3d.Point PointYmax0 = null;// là điểm có tọa độ X nhỏ hơn
        public t3d.Point PointYmax1 = null;// là điểm có tọa độ X lớn hơn
        public t3d.Point PointZmin = null;
        public t3d.Point PointZmax = null;
        public t3d.Point PointXminYmin = null;
        public t3d.Point PointXminYmax = null;
        public t3d.Point PointXmaxYmax = null;
        public t3d.Point PointXmaxYmin = null;
        public t3d.Point PointMidTop = null;
        public t3d.Point PointMidBot = null;
        public t3d.Point PointMidLeft = null;
        public t3d.Point PointMidRight = null;
        public List<t3d.Point> List_Edge { get; set; } // Danh sách tất cả các điểm edge của part
        public List<t3d.Point> List_ReferenceLinePoint { get; set; } // Danh sách tất cả các điểm edge của part
        //public List<t3d.Point> List_Edge_X { get; set; }//Danh sách 2 điểm Xmin và Xmax
        //public List<t3d.Point> List_Edge_Y { get; set; } //Danh sách 2 điểm Ymin và Ymax

        public Part_Edge(tsd.View viewcur, tsm.Part part) //Lấy các điểm part edge, center line, reference line của PART (trường hợp contour plate, poly beam chưa xét)
        {
            try
            {
                //tsm.WorkPlaneHandler wph = myModel.GetWorkPlaneHandler();
                frm_Main.wph.SetCurrentTransformationPlane(new TransformationPlane()); // CÓ CẦN BƯỚC CHUYỂN HỆ TỌA ĐỘ VỀ GỐC KHÔNG ?
                                                                                       //ArrayList List_RefLine = part.GetReferenceLine(false);
                                                                                       //MessageBox.Show(List_RefLine[0].ToString());
                                                                                       //foreach (object item in List_RefLine)
                                                                                       //{
                                                                                       //    MessageBox.Show(item.ToString());
                                                                                       //}
                tsd.ViewBase viewbase = viewcur as tsd.ViewBase;
                tsd.View view = viewbase as tsd.View;

                frm_Main.wph.SetCurrentTransformationPlane(new TransformationPlane(viewcur.DisplayCoordinateSystem)); //Đưa về tọa độ của view chứa drawing part
                                                                                                                      //frm_Main.model.CommitChanges();
                tsm.Solid partsolid = part.GetSolid();//Lấy Solic của part trên model
                tss.FaceEnumerator faceenum = partsolid.GetFaceEnumerator(); //Lấy tất các Face của part
                                                                             //Dùng While để duyệt qua từng face của part. Trường hợp này không dùng Forech được
                                                                             //Nếu dùng Foreach không báo lỗi thì dùng được

                //Khai báo danh sách các điểm của đối tượng (không trùng nhau)
                List<t3d.Point> PointList = new List<t3d.Point>();
                List<t3d.Point> PointListZ = new List<t3d.Point>(); //Chứa tất cả các điểm lấy được, để lọc ra lấy điểm có Zmin Zmax
                while (faceenum.MoveNext())
                {
                    tss.Face face = faceenum.Current as tss.Face;  //Convert current về face
                    tss.LoopEnumerator loopenum = face.GetLoopEnumerator(); //Gets a new loop enumerator for enumerating through the face's loops. 
                    while (loopenum.MoveNext())  //Duyệt  qua từng loop trong loopenum
                    {
                        //convert về loop
                        tss.Loop loop = loopenum.Current as tss.Loop;
                        tss.VertexEnumerator vertexenum = loop.GetVertexEnumerator();  //get vertexenum
                        while (vertexenum.MoveNext())   //Duyệt qua từng vertext
                        {
                            t3d.Point point = vertexenum.Current as t3d.Point;
                            PointListZ.Add(point);
                            //Kiểm tra điểm trùng với linq
                            t3d.Point diemtrung = PointList.Find(diem => Math.Abs(point.X - diem.X) < 0.01 && Math.Abs(point.Y - diem.Y) < 0.01);
                            if (diemtrung == null)
                            {
                                //MessageBox.Show(point.ToString());
                                PointList.Add(new t3d.Point(point.X, point.Y));
                                //tsd.Symbol symbol = new tsd.Symbol(viewbase, point, new tsd.SymbolInfo("xsteel", 1));
                                //symbol.Insert();
                            }
                        }
                        break;
                    }
                    //List<t3d.Point> minmaxX = new List<t3d.Point>();
                    t3d.Point pointZmin = PointListZ.OrderBy(point => point.Z).ToList()[0];          //Lấy minZ
                    t3d.Point pointZmax = PointListZ.OrderByDescending(point => point.Z).ToList()[0];//Lấy maxZ  
                    //t3d.Point pointXmin = PointList.OrderBy(point => point.X).ToList()[0];          //Lấy minX

                    List<t3d.Point> PointListOrderBy_X = PointList.OrderBy(point => point.X).ToList(); //để Lấy Xmin0 và Xmin1 (nếu có) 
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

                    List<t3d.Point> PointListOrderByDescending_X = PointList.OrderByDescending(point => point.X).ToList();//để Lấy Xmax0 và Xmax1 (nếu có) 
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

                    List<t3d.Point> PointListOrderBy_Y = PointList.OrderBy(point => point.Y).ToList();              //Lấy minY
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

                    List<t3d.Point> PointListOrderByDescending_Y = PointList.OrderByDescending(point => point.Y).ToList();    //Lấy maxY
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


                    PointZmin = pointZmin;
                    PointZmax = pointZmax;
                    PointXminYmin = new t3d.Point(pointXmin0.X, pointYmin0.Y);
                    PointXminYmax = new t3d.Point(pointXmin0.X, pointYmax0.Y);
                    PointXmaxYmax = new t3d.Point(pointXmax0.X, pointYmax0.Y);
                    PointXmaxYmin = new t3d.Point(pointXmax0.X, pointYmin0.Y);

                    PointMidTop = new t3d.Point((PointXminYmax.X + PointXmaxYmax.X) / 2, (PointXminYmax.Y + PointXmaxYmax.Y) / 2);         //tính lại
                    PointMidBot = new t3d.Point((PointXminYmin.X + PointXmaxYmin.X) / 2, (PointXminYmin.Y + PointXmaxYmin.Y) / 2);         //tính lại
                    PointMidLeft = new t3d.Point((PointXminYmin.X + PointXminYmax.X) / 2, (PointXminYmin.Y + PointXminYmax.Y) / 2);        //tính lại
                    PointMidRight = new t3d.Point((PointXmaxYmin.Y + PointXmaxYmax.Y) / 2, (PointXmaxYmin.X + PointXmaxYmax.X) / 2);       //tính lại

                    List_Edge = PointList;
                    //List_Edge_X = minmaxX; //Gồm 2 điểm min và max theo phương X
                    //List_Edge_Y = minmaxY; //Gồm 2 điểm min và max theo phương Y
                }
            }
            catch
            {
            }

        }
    }
}
