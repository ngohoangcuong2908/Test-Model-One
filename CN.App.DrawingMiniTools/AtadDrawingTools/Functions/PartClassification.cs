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
using Tekla.Structures.Datatype;
using tda = Tekla.Structures.Datatype;

using ATADDrawingTools.DataTypes;
using System.Text.RegularExpressions;

using System.Collections;
using System.IO;
using Newtonsoft.Json;
using ATADDrawingTools.Functions;


namespace ATADDrawingTools.Functions
{
    class PartClassification
    {
        public List<tsm.Beam> Webs { get; set; } = new List<tsm.Beam>();
        public List<tsm.Part> FlangeTops { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> FlangeBots { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> Stiffs { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> StiffsLeft { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> StiffsRight { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> Plates { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> PlatesLeft { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> PlatesRight { get; set; } = new List<tsm.Part>();
        public List<tsm.Part> HotRollPart { get; set; } = new List<tsm.Part>();

        public tsd.View View = null;

        /// <summary>
        /// Phân loại part theo vị trí so với gốc tọa độ của view
        /// Nếu plate có minZ và maxZ > 0 thì quy định nằm bên trái, ngược lại nằm bên phải
        /// Nếu plate không hoàn toàn nằm bên trái hoặc phải gốc tọa độ thì
        /// Nếu profile thep hình profile type != B thì cho vào list thép hình
        /// </summary>
        /// <param name="listPart"></param>
        public PartClassification(List<tsd.Part> listPart)
        {
            //tsd.Part drPart = new tsd.Part();

            //while (drobjenum.MoveNext())
            //{
            //    drPart = drobjenum.Current as tsd.Part; // convert drObj về drawing part 
            //    break;
            //}
            //tsd.View curview = drPart.GetView() as tsd.View;
            //tsd.ViewBase viewbase = drPart.GetView();
            //MessageBox.Show(drobjenum.GetSize().ToString());

            //frm_Main.wph.SetCurrentTransformationPlane(new TransformationPlane(curview.DisplayCoordinateSystem)); //Đưa về tọa độ của view chứa drawing part
            //frm_Main.model.CommitChanges();
            ////frm_Main.model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane(curview.DisplayCoordinateSystem)); //nếu để dòng lệnh này thì tọa độ sẽ nhảy lung tung. chưa biết tại sao
            //tsm.Part mPart = frm_Main.model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part; //lấy model part của drawing part
            //tsm.Part mainPart = mPart.GetAssembly().GetMainPart() as tsm.Part; //lấy mainpart của assemply chứa mPart
            //tsm.Assembly assembly = mPart.GetAssembly();
            //ArrayList secParts = assembly.GetSecondaries();
            //List<tsm.Part> secPartList = new List<tsm.Part>();
            //foreach (tsm.Part item in secParts) // vì GetSecondaries có thể lấy cả Sub ASS
            //{
            //    if (item is tsm.Part)
            //    {
            //        secPartList.Add(item);
            //    }
            //}
            foreach (tsd.Part drPart in listPart)
            {
                tsm.Part mPart = frm_Main.model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                tsd.View curview = drPart.GetView() as tsd.View;
                View = curview;
                Function function = new Function();
                double thickness = function.GetWidth(mPart);
                string partProfileType = function.GetProfileType(mPart);
                Part_Edge partEdge = new Part_Edge(curview, mPart);
                t3d.Point pMinX = partEdge.PointXmin0;
                t3d.Point pMaxX = partEdge.PointXmax0;
                t3d.Point pMinY = partEdge.PointYmin0;
                t3d.Point pMaxY = partEdge.PointYmax0;
                t3d.Point pMinZ = partEdge.PointZmin;
                t3d.Point pMaxZ = partEdge.PointZmax;
                //Kiem tra vi tri cua item
                if (Math.Abs(pMaxZ.Z - pMinZ.Z - thickness) < 3 && partProfileType =="B") continue; //điều kiện này đúng có nghĩa đây là tấm thấy bolt tròn
                bool isLeft = false;
                if (pMinZ.Z > 0 && partProfileType == "B") isLeft = true;
                if (isLeft)
                {
                    //MessageBox.Show("is left");
                    PlatesLeft.Add(mPart);
                }
                else if (!isLeft && partProfileType == "B")
                {
                   // MessageBox.Show("is right");
                    PlatesRight.Add(mPart);
                }
                else if (partProfileType != "B")
                {
                    HotRollPart.Add(mPart);
                }
            }
        }
    }
}
