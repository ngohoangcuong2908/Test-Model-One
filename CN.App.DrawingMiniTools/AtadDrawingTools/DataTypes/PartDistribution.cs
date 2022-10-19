using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
//Khai báo namespace của Tekla
using Tekla.Structures.Drawing;
using t3d = Tekla.Structures.Geometry3d;
using tsd = Tekla.Structures.Drawing;
//Khai báo shortcut cho các Namespace
using tsm = Tekla.Structures.Model;

namespace ATADDrawingTools.DataTypes
{
    class PartDistribution
    {
        //public string profileStringLetterOnly = string.Empty;
        public tsd.View View = null;
        public tsd.ViewBase ViewBase = null;
        //public string MainPartProfileType = string.Empty;
        public string PartProfileType = string.Empty;
        //public tsm.Part MainPart = null;
        public tsm.Part TopFlangeSection = null; //Cánh trên của thép built up trong section view, để phân biệt flange nằm ngang
        public tsm.Part BotFlangeSection = null;//Cánh dưới của thép built up trong section view

        public List<tsm.Part> ListEndPlatesLeft = new List<tsm.Part>(); //ds các tấm plate bên trái mainpart (có thê là tấm đứng hoặc nghiêng trái/ phải.
        public List<tsm.Part> ListEndPlatesRight = new List<tsm.Part>(); //ds các tấm plate bên phải mainpart (có thê là tấm đứng hoặc nghiêng trái/ phải.

        public List<tsm.Part> ListTamThayBoltTron = new List<tsm.Part>(); //ds các tấm plate thấy bolt tròn
        public List<tsm.Part> ListTamDung = new List<tsm.Part>(); //ds các tấm plate đứng
        public List<tsm.Part> ListTamNgang = new List<tsm.Part>();//ds các tấm plate nằm ngang
        public List<tsm.Part> ListTamNghiengTrai = new List<tsm.Part>();//ds các tấm plate nằm nghiêng trái
        public List<tsm.Part> ListTamNghiengPhai = new List<tsm.Part>();//ds các tấm plate nằm nghiêng phải

        public List<tsm.Part> ListPartINgang = new List<tsm.Part>();//ds các Part profile I nằm Ngang
        public List<tsm.Part> ListPartIDung = new List<tsm.Part>();//ds các Part profile I nằm Dung
        public List<tsm.Part> ListPartINghiengPhai = new List<tsm.Part>();//ds các Part profile I nằm nghiêng phải
        public List<tsm.Part> ListPartINghiengTrai = new List<tsm.Part>();//ds các Part profile I nằm nghiêng trai
        public List<tsm.Part> ListPartLUsection = new List<tsm.Part>();//ds các Part profile L U thấy section L U

        public bool havePart = false;
        public bool haveBolt = false;

        public List<List<t3d.Point>> listListPointsBolt_XY = new List<List<t3d.Point>>();
        public List<List<t3d.Point>> listListPointsBolt_Y = new List<List<t3d.Point>>();
        public List<List<t3d.Point>> listListPointsBolt_X = new List<List<t3d.Point>>();

        public List<tsm.Part> PlatesLeft { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> PlatesRight { get; set; } = new List<tsm.Part>();
        public PartDistribution(tsm.Part mainPart, tsd.DrawingObjectEnumerator drobjenum)
        {
            //string mainPartProfileType = null;
            string partProfileType = null;
            double chieuDayPlate = 0;
            foreach (tsd.DrawingObject drObj in drobjenum)
            {
                if (drObj is tsd.Part)
                {
                    havePart = true;
                    tsd.Part drPart = drObj as tsd.Part; // convert drObj về drawing part
                    tsd.View curview = drPart.GetView() as tsd.View;
                    tsd.ViewBase viewbase = drPart.GetView();
                    View = curview;
                    ViewBase = viewbase;
                    //frm_Main.model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(curview.DisplayCoordinateSystem)); //nếu để dòng lệnh này thì tọa độ sẽ nhảy lung tung. chưa biết tại sao
                    tsm.Part mPart = frm_Main.model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part; //lấy model part của drawing part
                                                                                                           //tsm.Part mainPart = mPart.GetAssembly().GetMainPart() as tsm.Part; //lấy mainpart của assemply chứa mPart
                    if (mPart.Identifier.GUID != mainPart.Identifier.GUID)
                    {
                        //Part_edge phải khởi tạo ở trên đây, trước khi lấy các điểm tính toán. Vì
                        // trong part_edge có phương thức đổi hệ tọa độ về hệ tọa độ bản vẽ, khi đó tọa độ điểm mới chính xác
                        Part_Edge part_edgeMainPart = new Part_Edge(View, mainPart);
                        Part_Edge part_edge = new Part_Edge(View, mPart);
                        ArrayList arrayListCenterLine = mPart.GetCenterLine(false);
                        List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                        t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                        t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                        //MessageBox.Show(pointCenterLineXmin.ToString() + "   " + pointCenterLineXmax.ToString());
                        mPart.GetReportProperty("PROFILE.WIDTH", ref chieuDayPlate);
                        mPart.GetReportProperty("PROFILE_TYPE", ref partProfileType);
                        PartProfileType = partProfileType; //đưa PartProfileType ra ngoài.
                                                           //Tính điểm định nghĩa của part theo bản vẽ

                        t3d.Point pointMiddleTopMainpart = part_edgeMainPart.PointMidTop; //điểm giữa cao nhất main part
                        t3d.Point pointMiddleBotMainpart = part_edgeMainPart.PointMidBot;//điểm giữa thấp nhất main part
                        frm_Main.model.CommitChanges();

                        //List<t3d.Point> part_edge_X = part_edge.List_Edge_X;
                        //List<t3d.Point> part_edge_Y = part_edge.List_Edge_Y;
                        t3d.Point minx = part_edge.PointXmin0;
                        t3d.Point maxx = part_edge.PointXmax0;
                        t3d.Point miny = part_edge.PointYmin0;
                        t3d.Point maxy = part_edge.PointYmax0;
                        t3d.Point minZ = part_edge.PointZmin;
                        t3d.Point maxZ = part_edge.PointZmax;
                        // MessageBox.Show(mPart.GetPartMark() + "-" + minZ.Z.ToString() + "-" + maxZ.Z.ToString() + " Restricbox Z: " + View.RestrictionBox.MinPoint.Z.ToString());
                        tsm.ModelObjectEnumerator boltOfPart = mPart.GetBolts();
                        if (minZ.Z >= View.RestrictionBox.MinPoint.Z - 3 && maxZ.Z <= View.RestrictionBox.MaxPoint.Z + 3) // điều kiện part nằm hoàn toàn trong view
                        {
                            //MessageBox.Show(maxx.ToString() + "_" + minx.ToString() + "_" + chieuDayPlate.ToString());
                            if (Math.Abs(maxx.X - minx.X - chieuDayPlate) < 3 && partProfileType == "B")// Tấm đứng. Bổ sung điều kiện kiểm tra vị trí so với main part để xem có phải endplate không?
                                ListTamDung.Add(mPart);
                            else if (Math.Abs(maxy.Y - miny.Y - chieuDayPlate) < 3 && partProfileType == "B")// Tấm ngang
                                ListTamNgang.Add(mPart);
                            else if (Math.Abs(t3d.Distance.PointToPoint(minx, miny) - chieuDayPlate) < 3 && partProfileType == "B") // Tấm nghiêng phải
                                ListTamNghiengPhai.Add(mPart);
                            //MessageBox.Show("Tam nghiêng phải" + mPart.GetPartMark().ToString());
                            else if (Math.Abs(t3d.Distance.PointToPoint(miny, maxx) - chieuDayPlate) < 3 && partProfileType == "B") // Tấm nghiêng trái 
                                ListTamNghiengTrai.Add(mPart);
                            //MessageBox.Show("Tam nghiêng trái" + mPart.GetPartMark().ToString());
                            else if (Math.Abs(t3d.Distance.PointToPoint(maxx, minx)) > 3 && Math.Abs(t3d.Distance.PointToPoint(maxy, miny)) > 3 && partProfileType == "B" && boltOfPart.GetSize() != 0)
                                ListTamThayBoltTron.Add(mPart);
                            else if (Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) < 2 && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 5)//Part I nằm NGANG     && partProfileType == "I"
                                ListPartINgang.Add(mPart);
                            else if (Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) < 2 && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 5) //Part I ĐỨNG          && partProfileType == "I"
                                ListPartIDung.Add(mPart);
                            else if (pointCenterLineXmin.Y > pointCenterLineXmax.Y && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 2 && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part I NGHIÊNG TRÁI 
                                ListPartINghiengTrai.Add(mPart);
                            else if (pointCenterLineXmin.Y < pointCenterLineXmax.Y && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 2 && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part I NGHIÊNG PHẢI
                                ListPartINghiengPhai.Add(mPart);
                            else if (Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) < 2 && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) < 2 && (partProfileType == "L" || partProfileType == "U")) //Part L U section
                                ListPartLUsection.Add(mPart);
                        }
                        else
                        {
                            if (mPart.PartNumber.Prefix == "F" && Math.Abs(miny.Y - pointMiddleTopMainpart.Y) < 3)
                            {
                                TopFlangeSection = mPart;
                                //MessageBox.Show("Top Flange: "+mPart.GetPartMark());
                            }
                            else if (mPart.PartNumber.Prefix == "F" && Math.Abs(maxy.Y - pointMiddleBotMainpart.Y) < 3)
                            {
                                BotFlangeSection = mPart;
                                //MessageBox.Show("Bot Flange: "+mPart.GetPartMark());
                            }

                        }
                    }
                }
                else if (drObj is tsd.Bolt)
                {
                    haveBolt = true;
                    //Convert về Bolt
                    tsd.Bolt drbolt = drObj as tsd.Bolt;
                    tsd.View curview = drbolt.GetView() as tsd.View;
                    Add_Mark add_mark = new Add_Mark(drbolt);//******************************************************************
                                                             //Lấy model BoltGroup thông qua part model identifỉe
                    tsm.BoltGroup mbolt = frm_Main.model.SelectModelObject(drbolt.ModelIdentifier) as tsm.BoltGroup;
                    Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(curview, mbolt);
                    //Tìm bolt theo phương XY
                    tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY; //Bolt xy là bolt thấy hình tròn.
                    tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y; //Bolt y là bolt nằm dọc.
                    tsm.BoltGroup bolt_x = tinhbolt.Bolt_X; //Bolt x là bolt nằm ngang.Z

                    List<t3d.Point> pointsBolt_XY = new List<t3d.Point>(); //danh sách các điểm của 1 nhóm bolt
                    List<t3d.Point> pointsBolt_X = new List<t3d.Point>(); //danh sách các điểm của 1 nhóm bolt
                    List<t3d.Point> pointsBolt_Y = new List<t3d.Point>(); //danh sách các điểm của 1 nhóm bolt

                    if (bolt_xy != null) //Nếu bolt xy không rỗng thì thực hiện các lệnh bên trong
                    {
                        bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                        foreach (t3d.Point p in bolt_xy.BoltPositions)
                        {
                            pointsBolt_XY.Add(p);
                        }
                        listListPointsBolt_XY.Add(pointsBolt_XY);
                    }


                    if (bolt_y != null) //Nếu bolt y(bolt nằm dọc) không rỗng thì thực hiện các lệnh bên trong
                    {
                        bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                        foreach (t3d.Point p in bolt_y.BoltPositions)
                        {
                            pointsBolt_Y.Add(p);
                        }
                        listListPointsBolt_Y.Add(pointsBolt_Y);
                    }
                    if (bolt_x != null) //Nếu bolt y(bolt nằm dọc) không rỗng thì thực hiện các lệnh bên trong
                    {
                        bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                        foreach (t3d.Point p in bolt_x.BoltPositions)
                        {
                            pointsBolt_X.Add(p);
                        }
                        listListPointsBolt_X.Add(pointsBolt_X);
                    }
                }
            }
        }

    }
}
