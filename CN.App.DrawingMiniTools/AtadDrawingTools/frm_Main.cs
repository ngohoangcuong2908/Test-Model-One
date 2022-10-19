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
using tsmui = Tekla.Structures.Model.UI;
using t3d = Tekla.Structures.Geometry3d;
using tsd = Tekla.Structures.Drawing;
using tsm = Tekla.Structures.Model;
using tsdui = Tekla.Structures.Drawing.UI;

using ATADDrawingTools.DataTypes;
using ATADDrawingTools.Functions;
using CN.App.DrawingTools.CopySectionDetail;
using System.Text.RegularExpressions;
using System.Collections;
using System.IO;
using Newtonsoft.Json;
using System.Diagnostics; //Để sử dụng Process mở file


namespace ATADDrawingTools //ApplicationFormBase
{
    public partial class frm_Main : Form // Đổi ntn để load được load save saveas...
    {
        //Khoi tao ket noi voi mo hinh
        public static tsm.Model model = new Model();
        public static tsd.DrawingHandler drawingHandler = new tsd.DrawingHandler(); //tao mot drawinghandler de co the tuong tac voi ban ve.
        public static tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();

        public static frm_Main _form;
        _ViewData input;
        _ViewData output;
        public frm_Main()
        {
            if (drawingHandler.GetConnectionStatus())
            {
                InitializeComponent();// Đổi ntn để load được load save saveas...
                _form = this;

                input = new _ViewData();
                output = new _ViewData();
            }
            else
            {
                MessageBox.Show("Tekla Structures must be opened!");
                this.Close();
                frm_Main.ActiveForm.Close();
            }
        }
        //Tao ra list chua danh sach ban ve 290, 180   281, 195   577, 509
        List<tsd.Drawing> drawingList = new List<Drawing>();
        public int minFormHeight = 79;
        public int maxFormHeight = 523;
        public int midFormHeight = 428;
        public int minFormWidth = 290;
        public int maxFormWidth = 643;

        public void frm_Main_Load(object sender, EventArgs e) // Các dòng lệnh trong này sẽ chạy ngay khi form được chạy.
        {
            try
            {
                DateTime expiryDate = new DateTime(2024, 07, 26);
                DateTime now = DateTime.Now;
                if (now > expiryDate)
                {
                    frm_Main.ActiveForm.Close();
                }
                //Lay duong dan cua model dang duoc mo
                string modelPath = model.GetInfo().ModelPath;
                string attributePath = Path.Combine(modelPath, "attributes");
                //Lay cac file o trong attribute folder
                string[] drAsAttributes = Directory.GetFiles(attributePath, "*.ad"); //load attribute bản vẽ assembly trong folder Model
                string[] drSgAttributes = Directory.GetFiles(attributePath, "*.wd"); //load attribute bản vẽ assembly trong folder Model
                string[] drDimAttributes = Directory.GetFiles(attributePath, "*.dim"); //load attribute DIM trong folder Model
                string[] drSectionMarkAttributes = Directory.GetFiles(attributePath, "*.cs"); //load attribute DIM trong folder Model
                string[] drSectionViewAttributes = Directory.GetFiles(attributePath, "*.vi"); //load attribute DIM trong folder Model
                if (modelPath.Last() != '\\')
                    modelPath = modelPath + '\\';
                tbx_DirectoryOpenDrawing.Text = modelPath; //Hiển thị đường dẫn model ở ô textbox tab Numbering drawing, để mở drawing active
                //Tạo ra 1 list phụ để kiểm tra phần tử trùng
                List<string> singlePartAttributes = new List<string>();
                List<string> assemblyAttributes = new List<string>();
                foreach (var item in drDimAttributes)
                {
                    cb_DimentionAttributes.Items.Add(Path.GetFileNameWithoutExtension(item)); //đưa dữ liệu vào ô DimentionProperties.
                    cb_DimAttDimPartBoltAssembly.Items.Add(Path.GetFileNameWithoutExtension(item)); //đưa dữ liệu vào ô cb_DimAttDimPartBoltAssembly của button dim parts/bolts of assembly
                }
                foreach (var item in drSectionMarkAttributes)
                {
                    cb_SecMarkAttribute.Items.Add(Path.GetFileNameWithoutExtension(item)); //đưa dữ liệu vào ô cb_SectionMarkAttributes.
                }
                foreach (var item in drSectionViewAttributes)
                {
                    cb_SecAttribute.Items.Add(Path.GetFileNameWithoutExtension(item)); //đưa dữ liệu vào ô cb_SectionMarkAttributes.
                }
                this.KeyPreview = true; // để tạo phím tắt cho các nút
                FormLoadDefault();
                GetDrawingInfor();
                //txt_RevisionDrawing.Text = CN.App.DrawingTools.Properties.Settings.Default.Name;
                //cb_SecMarkAttribute.SelectedItem = "";
                //cb_DimentionAttributes.SelectedItem = drDimAttributes.Length - 1;
                //cb_DimAttDimPartBoltAssembly.SelectedItem = 3;
            }
            //((DataGridViewComboBoxColumn)(dtgv_AssDrawingAttribute.Columns[dt_AssDrAttribute.Index])).DataSource = cb_ASDrAttributes.Items;
            catch { }
        }
        private void chbox_Topmost_CheckedChanged(object sender, EventArgs e)
        {
            if (chbox_Topmost.Checked)
                this.TopMost = true;
            else this.TopMost = false;
        }
        private void tb_FormOpacity_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                double opacity = Convert.ToDouble(tb_FormOpacity.Text);
                if (opacity >= 10)
                    this.Opacity = Convert.ToDouble(tb_FormOpacity.Text) * 0.01;
                else this.Opacity = 10 * 0.01;
            }
        }
        private void FormLoadDefault()
        {
            frm_Main.ActiveForm.Width = CN.App.DrawingMiniTools.Properties.Settings.Default.FormWidth_df;
            frm_Main.ActiveForm.Height = CN.App.DrawingMiniTools.Properties.Settings.Default.FormHeight_df;
            frm_Main.ActiveForm.Location = new System.Drawing.Point(CN.App.DrawingMiniTools.Properties.Settings.Default.FormLocationX, CN.App.DrawingMiniTools.Properties.Settings.Default.FormLocationY);
            frm_Main.ActiveForm.TopMost = true;

            cb_DimAttDimPartBoltAssembly.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.cb_DimAttDimPartBoltAssembly_df;
            cb_PartBoltOption.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.cb_PartBoltOption_df;
            chbox_DimInternalPart.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_DimInternalPart_df;
            chbox_DimClose.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_DimClose_df;
            cb_SecMarkAttribute.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.cb_SecMarkAttribute_df;
            cb_SecAttribute.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.cb_SecAttribute_df;
            tb_SectionMarkPrefix.Text = CN.App.DrawingMiniTools.Properties.Settings.Default.tb_SectionMarkPrefix_df;
            tb_SectionMarkNo.Text = CN.App.DrawingMiniTools.Properties.Settings.Default.tb_SectionMarkNo_df;
            chbox_DimInternalPartInAssembly.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_DimInternalPartInAssembly_df;
            imcb_SelectDimRefCenEdge.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.imcb_SelectDimRefCenEdge_df;
            cb_DimentionAttributes.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.cb_DimentionAttributes_df;
            imcb_DimensionTo.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.imcb_DimensionTo_df;
            txt_RevText.Text = CN.App.DrawingMiniTools.Properties.Settings.Default.txt_RevText_df;
            txt_RevisionDrawing.Text = CN.App.DrawingMiniTools.Properties.Settings.Default.txt_RevisionDrawing_df;
            cb_SelectTitleReviseText.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.cb_SelectTitleReviseText_df;
            tb_DimSpaceDim2Part.Text = CN.App.DrawingMiniTools.Properties.Settings.Default.tb_DimSpaceDim2Part_df;
            tb_DimSpaceDim2Dim.Text = CN.App.DrawingMiniTools.Properties.Settings.Default.tb_DimSpaceDim2Dim_df;
            chbox_AddPartMark.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AddPartMark_df;
            chbox_ClearCurrentDim.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_ClearCurrentDim_df;
            imcb_SectionSide.SelectedIndex = CN.App.DrawingMiniTools.Properties.Settings.Default.imcb_SectionSide_df;
            chbox_SkewBolts.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_SkewBolts_df;
            chbox_AutoCreateSection.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AutoCreateSection_df;
            chbox_AutoFixCutPartLength.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AutoFixCutPartLength_df;
            chbox_AddLine.Checked = CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AddLine_df;
        }

        private void frm_Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            CN.App.DrawingMiniTools.Properties.Settings.Default.FormLocationX = this.Location.X;
            CN.App.DrawingMiniTools.Properties.Settings.Default.FormLocationY = this.Location.Y;
            CN.App.DrawingMiniTools.Properties.Settings.Default.FormWidth_df = this.Width;
            CN.App.DrawingMiniTools.Properties.Settings.Default.FormHeight_df = this.Height;
            CN.App.DrawingMiniTools.Properties.Settings.Default.cb_DimAttDimPartBoltAssembly_df = cb_DimAttDimPartBoltAssembly.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.cb_PartBoltOption_df = cb_PartBoltOption.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_DimInternalPart_df = chbox_DimInternalPart.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_DimClose_df = chbox_DimClose.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.cb_SecMarkAttribute_df = cb_SecMarkAttribute.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.cb_SecAttribute_df = cb_SecAttribute.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.tb_SectionMarkPrefix_df = tb_SectionMarkPrefix.Text;
            CN.App.DrawingMiniTools.Properties.Settings.Default.tb_SectionMarkNo_df = tb_SectionMarkNo.Text;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_DimInternalPartInAssembly_df = chbox_DimInternalPartInAssembly.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.imcb_SelectDimRefCenEdge_df = imcb_SelectDimRefCenEdge.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.cb_DimentionAttributes_df = cb_DimentionAttributes.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.imcb_DimensionTo_df = imcb_DimensionTo.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.txt_RevText_df = txt_RevText.Text;
            CN.App.DrawingMiniTools.Properties.Settings.Default.txt_RevisionDrawing_df = txt_RevisionDrawing.Text;
            CN.App.DrawingMiniTools.Properties.Settings.Default.cb_SelectTitleReviseText_df = cb_SelectTitleReviseText.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.tb_DimSpaceDim2Part_df = tb_DimSpaceDim2Part.Text;
            CN.App.DrawingMiniTools.Properties.Settings.Default.tb_DimSpaceDim2Dim_df = tb_DimSpaceDim2Dim.Text;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AddPartMark_df = chbox_AddPartMark.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_ClearCurrentDim_df = chbox_ClearCurrentDim.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.imcb_SectionSide_df = imcb_SectionSide.SelectedIndex;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_SkewBolts_df = chbox_SkewBolts.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AutoCreateSection_df = chbox_AutoCreateSection.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AutoFixCutPartLength_df = chbox_AutoFixCutPartLength.Checked;
            CN.App.DrawingMiniTools.Properties.Settings.Default.chbox_AddLine_df = chbox_AddLine.Checked;

            CN.App.DrawingMiniTools.Properties.Settings.Default.Save();
        }

        public void GetDrawingInfor()
        {
            try
            {
                string lastRevDr = string.Empty;
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (cur_dr is tsd.AssemblyDrawing)
                {
                    tsd.AssemblyDrawing assemblyDrawing = cur_dr as tsd.AssemblyDrawing;
                    tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                    assembly.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                }
                else if (cur_dr is tsd.SinglePartDrawing)
                {
                    tsd.SinglePartDrawing singleDrawing = cur_dr as tsd.SinglePartDrawing;
                    tsm.Part part = model.SelectModelObject(singleDrawing.PartIdentifier) as tsm.Part;
                    part.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                }

                this.Text = cur_dr.Mark + " - R" + lastRevDr;
                tb_DrawingName.Text = cur_dr.Name;
                tb_DrawingTitle1.Text = cur_dr.Title1;
                tb_DrawingTitle2.Text = cur_dr.Title2;
                tb_DrawingTitle3.Text = cur_dr.Title3;
                bool isFrozen = cur_dr.IsFrozen;
                bool isIssued = cur_dr.IsIssued;
                bool isReadyForIssue = cur_dr.IsReadyForIssue;
                if (cur_dr.IsFrozen) btn_Frozen.BackColor = System.Drawing.Color.DeepSkyBlue;
                else btn_Frozen.BackColor = System.Drawing.Color.Silver;
                if (cur_dr.IsIssued) btn_Issue.BackColor = System.Drawing.Color.DarkOrange;
                else btn_Issue.BackColor = System.Drawing.Color.Silver;
                if (cur_dr.IsReadyForIssue) btn_Ready.BackColor = System.Drawing.Color.Green;
                else btn_Ready.BackColor = System.Drawing.Color.Silver;

                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetActiveDrawing().GetSheet().GetAllObjects(new Type[] { typeof(tsd.View) });
                foreach (tsd.View view in drobjenum)
                {
                    if (view.ViewType == tsd.View.ViewTypes.FrontView)
                    {
                        lb_MinCutPartLength_Scale.Text = "CP + SC: " + view.Attributes.Shortening.MinimumLength.ToString() + " - " + view.Attributes.Scale.ToString();
                        break;
                    }
                    else if (view.ViewType == tsd.View.ViewTypes.TopView)
                    {
                        lb_MinCutPartLength_Scale.Text = "CP + SC: " + view.Attributes.Shortening.MinimumLength.ToString() + " - " + view.Attributes.Scale.ToString();
                        break;
                    }
                    else if (view.ViewType == tsd.View.ViewTypes.BackView)
                    {
                        lb_MinCutPartLength_Scale.Text = "CP + SC: " + view.Attributes.Shortening.MinimumLength.ToString() + " - " + view.Attributes.Scale.ToString();
                        break;
                    }
                    else if (view.ViewType == tsd.View.ViewTypes.BottomView)
                    {
                        lb_MinCutPartLength_Scale.Text = "CP + SC: " + view.Attributes.Shortening.MinimumLength.ToString() + " - " + view.Attributes.Scale.ToString();
                        break;
                    }
                }

            }
            catch
            {
                this.Text = "No drawings open!";
                timer1.Enabled = false;
                //tb_DrawingName.Text = "";
                //tb_DrawingTitle1.Text = "";
                //tb_DrawingTitle2.Text = "";
                //tb_DrawingTitle3.Text = "";
                //btn_Frozen.BackColor = Color.Silver;
                //btn_Issue.BackColor = Color.Silver;
                //btn_Ready.BackColor = Color.Silver;
            }
        }
        private void frm_Main_Deactivate(object sender, EventArgs e) // dùng cho phần thông tin bản vẽ
        {
            try
            {
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (tb_DrawingName.Modified)
                {
                    cur_dr.Name = tb_DrawingName.Text;
                    cur_dr.Modify();
                }
                if (tb_DrawingTitle1.Modified)
                {
                    cur_dr.Title1 = tb_DrawingTitle1.Text;
                    cur_dr.Modify();
                }
                if (tb_DrawingTitle2.Modified)
                {
                    cur_dr.Title2 = tb_DrawingTitle2.Text;
                    cur_dr.Modify();
                }
                if (tb_DrawingTitle3.Modified)
                {
                    cur_dr.Title3 = tb_DrawingTitle3.Text;
                    cur_dr.Modify();
                }
                timer1.Enabled = true;
            }
            catch
            { timer1.Enabled = false; }
        }
        private void frm_Main_KeyDown(object sender, KeyEventArgs e) // cài đặt phím tắt của chương trình
        {
            if (e.KeyCode == Keys.D2 && e.Modifiers == Keys.Control)
            {
                OpenNextDrawing();
            }
            else if (e.KeyCode == Keys.D1 && e.Modifiers == Keys.Control)
            {
                OpenPreviousDrawing();
            }
            else if (e.KeyCode == Keys.F5)
            {
                Application.Restart();
                Environment.Exit(0);
            }
            else if (e.KeyCode == Keys.D0 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 0;
            }
            else if (e.KeyCode == Keys.D1 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 1;
            }
            else if (e.KeyCode == Keys.D2 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 2;
            }
            else if (e.KeyCode == Keys.D3 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 3;
            }
            else if (e.KeyCode == Keys.D4 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 4;
            }
            else if (e.KeyCode == Keys.D5 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 5;
            }
            else if (e.KeyCode == Keys.D6 && e.Modifiers == Keys.Alt)
            {
                cb_PartBoltOption.SelectedIndex = 6;
            }
        }
        private void cb_PartBoltOption_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_PartBoltOption.SelectedIndex == 1)
            {
                chbox_SkewBolts.Enabled = true;
                chbox_AutoFixCutPartLength.Enabled = false;
                chbox_AutoCreateSection.Enabled = false;
                chbox_AddLine.Enabled = false;
            }
            else if (cb_PartBoltOption.SelectedIndex == 5)
            {
                chbox_SkewBolts.Enabled = false;
                chbox_AutoFixCutPartLength.Enabled = true;
                chbox_AutoCreateSection.Enabled = true;
                chbox_AddLine.Enabled = true;
            }
            else
            {
                chbox_SkewBolts.Enabled = false;
                chbox_AutoFixCutPartLength.Enabled = false;
                chbox_AutoCreateSection.Enabled = false;
                chbox_AddLine.Enabled = false;
            }
        }

        #region Đoạn code này để kéo thả form được

        //public const int WM_NCLBUTTONDOWN = 0xA1;
        //public const int HT_CAPTION = 0x2;

        //[System.Runtime.InteropServices.DllImport("user32.dll")]
        //public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        //[System.Runtime.InteropServices.DllImport("user32.dll")]
        //public static extern bool ReleaseCapture();
        //private void frm_Main_MouseDown(object sender, MouseEventArgs e)
        //{
        //    if (e.Button == MouseButtons.Left)
        //    {
        //        ReleaseCapture();
        //        SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        //    }
        //}

        private void frm_Main_MouseDoubleClick(object sender, MouseEventArgs e) // Clik double left mouse để Minimized form
        {
            if (e.Button == MouseButtons.Left) // MouseButtons.Middle
            {
                this.WindowState = FormWindowState.Minimized;
            }
        }
        private void frm_Main_MouseClick(object sender, MouseEventArgs e) // Click chuột phải để mở rộng/ thu nhỏ form
        {
            if (e.Button == MouseButtons.Right)
            {
                if (!gb9)
                {
                    frm_Main.ActiveForm.Width = maxFormWidth;
                    frm_Main.ActiveForm.Height = maxFormHeight;
                }
                else
                {
                    frm_Main.ActiveForm.Width = minFormWidth;
                    frm_Main.ActiveForm.Height = minFormHeight;
                }
                gb9 = !gb9;
            }
        }
        #endregion


        // TAB ***********[DIMENSION TOOL]***********

        private void btn_RunDimGA_Click(object sender, EventArgs e)
        {/*
            if (imcb_SelectDimRefCenEdge.SelectedIndex == 0)  // Khi chon dim cho Reference line
            {
                try
                {
                    tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                    wph.SetCurrentTransformationPlane(new TransformationPlane());
                    tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                    // Lấy đối tượng được chọn trên bản vẽ
                    tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                    CreateDimension createDimension = new CreateDimension(); // tạo 1 createDimension xem thêm class CreateDimension
                    tsdui.Picker c_picker = drawingHandler.GetPicker();
                    var pick = c_picker.PickPoint("Chon diem thu nhat");
                    //Lấy tọa độ điểm pick (điểm gốc)
                    t3d.Point point1 = pick.Item1;
                    //Lấy view hiện hành thông qua điểm vừa pick
                    tsd.View viewcur = pick.Item2 as tsd.View;
                    tsd.ViewBase viewbase = pick.Item2;
                    tsd.PointList pointListDim = new tsd.PointList() { point1 };
                    try
                    {
                        var pick2 = c_picker.PickPoint("Chon diem thu hai");
                        t3d.Point point2 = pick2.Item1;
                        pointListDim.Add(point2);
                    }
                    catch
                    { }

                    foreach (tsd.DrawingObject drobj in drobjenum)
                    {
                        if (drobj is tsd.Part)
                        {
                            //Convert về Part
                            tsd.Part drpart = drobj as tsd.Part;
                            //Lấy model part thông qua part model identifỉe
                            tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                            tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                            wph.SetCurrentTransformationPlane(viewplane);
                            ArrayList listRefLine = mpart.GetReferenceLine(false);
                            foreach (t3d.Point p in listRefLine)
                            {
                                pointListDim.Add(p);
                            }
                            wph.SetCurrentTransformationPlane(new TransformationPlane());
                        }
                    }
                    if (imcb_DirectionXYFree.SelectedIndex == -1 || imcb_DirectionXYFree.SelectedIndex == 0) //DIM X
                    {
                        var pick3 = c_picker.PickPoint("Chon diem dat dim");
                        t3d.Point point3 = pick3.Item1;
                        if (point1.Y < point3.Y)
                        {
                            createDimension.CreateStraightDimensionSet_X(viewbase, pointListDim, cb_DimentionAttributes.Text, 250);
                        }
                        else
                        {
                            createDimension.CreateStraightDimensionSet_X_(viewbase, pointListDim, cb_DimentionAttributes.Text, 250);
                        }
                    }
                    if (imcb_DirectionXYFree.SelectedIndex == 1) //DIM Y
                    {
                        var pick3 = c_picker.PickPoint("Chon diem dat dim");
                        t3d.Point point3 = pick3.Item1;
                        if (point1.X < point3.X)
                        {
                            createDimension.CreateStraightDimensionSet_Y(viewbase, pointListDim, cb_DimentionAttributes.Text, 250);
                        }
                        else
                        {
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, pointListDim, cb_DimentionAttributes.Text, 250);
                        }
                    }
                    if (imcb_DirectionXYFree.SelectedIndex == 2) //DIM FREE
                    {
                        var pick4 = c_picker.PickPoint("Chon diem 1 xac dinh phuong");
                        t3d.Point point4 = pick4.Item1;
                        var pick5 = c_picker.PickPoint("Chon diem 2 xac dinh phuong");
                        t3d.Point point5 = pick5.Item1;
                        createDimension.CreateStraightDimensionSet_FX(point4, point5, viewbase, pointListDim, cb_DimentionAttributes.Text, 250);
                    }
                }
                catch
                { }
            }
            else if (imcb_SelectDimRefCenEdge.SelectedIndex == 1) // Khi chon dim cho center line
            {
                try
                {
                    tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                    wph.SetCurrentTransformationPlane(new TransformationPlane());
                    tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                    // Lấy đối tượng được chọn trên bản vẽ
                    tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                    CreateDimension createDimension = new CreateDimension(); // tạo 1 createDimension xem thêm class CreateDimension
                    tsdui.Picker c_picker = drawingHandler.GetPicker();
                    var pick = c_picker.PickPoint("Chon diem thu nhat");
                    //Lấy tọa độ điểm pick (điểm gốc)
                    t3d.Point point1 = pick.Item1;
                    //Lấy view hiện hành thông qua điểm vừa pick
                    tsd.View viewcur = pick.Item2 as tsd.View;
                    tsd.ViewBase viewbase = pick.Item2;
                    tsd.PointList pointListDim = new tsd.PointList() { point1 };
                    try
                    {
                        var pick2 = c_picker.PickPoint("Chon diem thu hai");
                        t3d.Point point2 = pick2.Item1;
                        pointListDim.Add(point2);
                    }
                    catch
                    { }


                    int distanceY = new int();
                    int distanceY_ = new int();
                    int distanceX = new int();
                    int distanceX_ = new int();
                    foreach (tsd.DrawingObject drobj in drobjenum)
                    {
                        if (drobj is tsd.Part)
                        {
                            //Convert về Part
                            tsd.Part drpart = drobj as tsd.Part;
                            //Lấy model part thông qua part model identifỉe
                            tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;

                            tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                            wph.SetCurrentTransformationPlane(viewplane);
                            ArrayList listCenterLine = mpart.GetCenterLine(false);
                            foreach (t3d.Point p in listCenterLine)
                            {
                                pointListDim.Add(p);
                            }
                            wph.SetCurrentTransformationPlane(new TransformationPlane());
                        }
                    }
                    if (imcb_DirectionXYFree.SelectedIndex == -1 || imcb_DirectionXYFree.SelectedIndex == 0)
                    {
                        var pick3 = c_picker.PickPoint("Chon diem dat dim");
                        t3d.Point point3 = pick3.Item1;
                        if (point1.Y < point3.Y)
                        {
                            distanceY = Convert.ToInt32(Math.Abs(point3.Y - pointListDim[pointListDim.Count - 1].Y));
                            createDimension.CreateStraightDimensionSet_X(viewbase, pointListDim, cb_DimentionAttributes.Text, distanceY);
                        }
                        else
                        {
                            distanceY_ = Convert.ToInt32(Math.Abs(point3.Y - pointListDim[0].Y));
                            createDimension.CreateStraightDimensionSet_X_(viewbase, pointListDim, cb_DimentionAttributes.Text, distanceY_);
                        }
                    }
                    if (imcb_DirectionXYFree.SelectedIndex == 1)
                    {
                        var pick3 = c_picker.PickPoint("Chon diem dat dim");
                        t3d.Point point3 = pick3.Item1;
                        if (point1.X < point3.X)
                        {
                            createDimension.CreateStraightDimensionSet_Y(viewbase, pointListDim, cb_DimentionAttributes.Text, distanceX);
                        }
                        else
                        {
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, pointListDim, cb_DimentionAttributes.Text, distanceX_);
                        }
                    }
                    if (imcb_DirectionXYFree.SelectedIndex == 2)
                    {
                        var pick4 = c_picker.PickPoint("Chon diem 1 xac dinh phuong");
                        t3d.Point point4 = pick4.Item1;
                        var pick5 = c_picker.PickPoint("Chon diem 2 xac dinh phuong");
                        t3d.Point point5 = pick5.Item1;
                        createDimension.CreateStraightDimensionSet_FX(point4, point5, viewbase, pointListDim, cb_DimentionAttributes.Text, distanceX);
                    }
                }
                catch
                { }
            }
            else if (imcb_SelectDimRefCenEdge.SelectedIndex == 2)  // Khi chon dim cho EDGE
            {
                try
                {
                    // Khai báo myDrawingHandler và model đã làm rồi (pulic static), nếu chưa có có thể thêm ở đây.
                    //Khai báo workplanhandler
                    tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                    wph.SetCurrentTransformationPlane(new TransformationPlane());
                    // tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                    // Lấy đối tượng được chọn trên bản vẽ
                    tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    //tsd.StraightDimensionSetHandler smyDrawingHandler = new StraightDimensionSetHandler();
                    CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                                                                             //tsd.DrawingHandler MyDrHd = new tsd.DrawingHandler();
                    tsdui.Picker c_picker = drawingHandler.GetPicker();
                    var pick = c_picker.PickPoint("Chon diem thu nhat");
                    //Lấy tọa độ điểm pick (điểm gốc)
                    t3d.Point point1 = pick.Item1;
                    t3d.Point point2 = null;
                    try
                    {
                        var pick2 = c_picker.PickPoint("Chon diem thu hai");
                        point2 = pick2.Item1;
                    }
                    catch
                    { }
                    var pick3 = c_picker.PickPoint("Chon diem dat dim");
                    t3d.Point point3 = pick3.Item1;
                    //Lấy view hiện hành thông qua điểm vừa pick
                    tsd.View viewcur = pick.Item2 as tsd.View;
                    tsd.ViewBase viewbase = pick.Item2;

                    tsd.PointList dimPartListX = new PointList();//Tập hợp điểm X bên trên
                    tsd.PointList dimPartListX_ = new PointList();//Tập hợp điểm X bên dưới
                    tsd.PointList dimPartListY = new PointList();//Tập hợp điểm Y bên phải
                    tsd.PointList dimPartListY_ = new PointList(); //Tập hợp điểm Y bên trái
                    dimPartListX.Add(point1);
                    dimPartListX_.Add(point1);
                    dimPartListY.Add(point1);
                    dimPartListY_.Add(point1);

                    int distanceY = new int();
                    int distanceY_ = new int();
                    int distanceX = new int();
                    int distanceX_ = new int();
                    foreach (tsd.DrawingObject drobj in drobjenum)
                    {
                        if (drobj is tsd.Part)
                        {
                            wph.SetCurrentTransformationPlane(new TransformationPlane());
                            //Convert về Part
                            tsd.Part drPart = drobj as tsd.Part;
                            //Lấy model part thông qua part model identifỉe
                            tsm.Part mPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                            //Lấy view hiện hành thông qua drpart
                            tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                            wph.SetCurrentTransformationPlane(viewplane);
                            Part_Edge part_edge = new Part_Edge(viewcur, mPart);
                            t3d.Point maxX = part_edge.PointXmax;
                            t3d.Point minX = part_edge.PointXmin;
                            t3d.Point maxY = part_edge.PointYmax;
                            t3d.Point minY = part_edge.PointYmin;

                            if (imcb_DirectionXYFree.SelectedIndex == -1 || imcb_DirectionXYFree.SelectedIndex == 0) //Dim X
                            {
                                dimPartListX.Add(part_edge.PointXminYmax);
                                dimPartListX.Add(part_edge.PointXmaxYmax);
                                dimPartListX_.Add(part_edge.PointXminYmin);
                                dimPartListX_.Add(part_edge.PointXmaxYmin);
                            }
                            if (imcb_DirectionXYFree.SelectedIndex == 1) //Dim Y
                            {
                                dimPartListY.Add(part_edge.PointXmaxYmax);
                                dimPartListY.Add(part_edge.PointXmaxYmin);
                                dimPartListY_.Add(part_edge.PointXminYmax);
                                dimPartListY_.Add(part_edge.PointXminYmin);
                            }
                        }
                    }
                    if (point2 != null)
                    {
                        dimPartListX.Add(point2);
                        dimPartListX_.Add(point2);
                        dimPartListY.Add(point2);
                        dimPartListY_.Add(point2);
                    }

                    if (imcb_DirectionXYFree.SelectedIndex == -1 || imcb_DirectionXYFree.SelectedIndex == 0)
                    {
                        if (point3.Y > dimPartListX[0].Y)
                        {
                            distanceY = Convert.ToInt32(Math.Abs(point3.Y - dimPartListX[dimPartListX.Count - 1].Y));
                            createDimension.CreateStraightDimensionSet_X(viewbase, dimPartListX, cb_DimentionAttributes.Text, distanceY); // tạo DIM part phương X

                            // tsd.StraightDimensionSet Dim_Bolt_X = smyDrawingHandler.CreateDimensionSet(viewbase, dimPartListX, new t3d.Vector(0, 1, 0), 250);
                        }
                        else
                        {
                            distanceY_ = Convert.ToInt32(Math.Abs(point3.Y - dimPartListX_[1].Y));
                            createDimension.CreateStraightDimensionSet_X_(viewbase, dimPartListX_, cb_DimentionAttributes.Text, distanceY_); // tạo DIM part phương -X
                                                                                                                                             //tsd.StraightDimensionSet Dim_Bolt_X = smyDrawingHandler.CreateDimensionSet(viewbase, dimPartListX_, new t3d.Vector(0, -1, 0), 250);
                        }
                    }

                    if (imcb_DirectionXYFree.SelectedIndex == 1)
                    {
                        if (point3.X > dimPartListY[0].X)
                        {
                            distanceX = Convert.ToInt32(Math.Abs(point3.X - dimPartListY[dimPartListY.Count - 1].X));
                            createDimension.CreateStraightDimensionSet_Y(viewbase, dimPartListY, cb_DimentionAttributes.Text, distanceX); // tạo DIM part phương Y
                                                                                                                                          //tsd.StraightDimensionSet Dim_Bolt_X = smyDrawingHandler.CreateDimensionSet(viewbase, dimPartListY, new t3d.Vector(1, 0, 0), 250);
                        }
                        else
                        {
                            distanceX_ = Convert.ToInt32(Math.Abs(point3.X - dimPartListY[0].X));
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, dimPartListY, cb_DimentionAttributes.Text, distanceX_); // tạo DIM part phương -Y
                                                                                                                                            // tsd.StraightDimensionSet Dim_Bolt_X = smyDrawingHandler.CreateDimensionSet(viewbase, dimPartListY_, new t3d.Vector(-1, 0, 0), 250);
                        }
                    }
                }
                catch
                { }
            }
            */
        } // SAVE CODE ONLY

        private void btn_DimPartOverallBolt_Click(object sender, EventArgs e)
        {
            try
            {
                //Nếu chọn dim grating thì gọi hàm DimGratingCheckerPlate
                if (cb_PartBoltOption.SelectedIndex == 2)
                {
                    DimGratingCheckerPlate();
                    goto Ketthuc;
                }
                //Nếu chọn dim viền thì gọi hàm DimTrimFlashing
                else if (cb_PartBoltOption.SelectedIndex == 3)
                {
                    DimTrimFlashing();
                    goto Ketthuc;
                }
                //Nếu check vào ô dim Curve tole
                else if (cb_PartBoltOption.SelectedIndex == 4)
                {
                    DimCurveTole();
                    goto Ketthuc;
                }
                //Nếu check vào ô dim Purlin
                else if (cb_PartBoltOption.SelectedIndex == 5)
                {
                    DimPurlin();
                    goto Ketthuc;
                }
                //Nếu check vào ô dim Arrestor
                else if (cb_PartBoltOption.SelectedIndex == 6)
                {
                    DimArrestor();
                    goto Ketthuc;
                }
                // Khai báo drawingHandler và model đã làm rồi (pulic static), nếu chưa có có thể thêm ở đây.
                //Khai báo workplanhandler
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                ClearDrawingObjects clearDrawingObjects = new ClearDrawingObjects();
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                DrawingObjectEnumerator.AutoFetch = true;
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsdui.Picker c_picker = drawingHandler.GetPicker();
                string dimAttribute = cb_DimentionAttributes.Text; //Khoảng cb_DimentionAttributes mặc định
                int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drPart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                        //Lấy chiều dày và profile của tấm
                        double chieuDayPlate = 0;
                        string partProfileType = null;
                        mPart.GetReportProperty("PROFILE.WIDTH", ref chieuDayPlate);
                        mPart.GetReportProperty("PROFILE_TYPE", ref partProfileType);

                        //Lấy view hiện hành thông qua drpart
                        tsd.View viewcur = drPart.GetView() as tsd.View;
                        tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                        double viewScale = viewcur.Attributes.Scale;
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.ViewBase viewBase = drPart.GetView();
                        if (chbox_ClearCurrentDim.Checked)
                            clearDrawingObjects.ClearDim(viewcur);
                        if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                        {
                            dimSpace = Convert.ToInt32(viewScale * 10);
                            dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                        }

                        //Tính điểm định nghĩa của part theo bản vẽ
                        Part_Edge partEdge = new Part_Edge(viewcur, mPart);

                        t3d.Point minx = partEdge.PointXmin0;
                        t3d.Point maxx = partEdge.PointXmax0;
                        t3d.Point miny = partEdge.PointYmin0;
                        t3d.Point maxy = partEdge.PointYmax0;

                        //MessageBox.Show(string.Format("minX {0} maxX {1} minY {2} maxY {3} thickness {4} xxx {5}", minx, maxx, miny, maxy, chieuDayPlate, Math.Abs(t3d.Distance.PointToPoint(minx, miny) - chieuDayPlate)));

                        if (Math.Abs(maxx.X - minx.X - chieuDayPlate) < 3 && partProfileType == "B")// Tấm đứng (front view), profile Plate
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();//Ds 2 điểm thấp nhất và cao nhất bên Phai

                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                pl.Add(new t3d.Point(partEdge.PointXminYmin));//add thêm Điểm có tọa độ X lon nhất và tọa độ Y nhỏ nhất
                                pl.Add(new t3d.Point(partEdge.PointXminYmax));//add thêm Điểm có tọa độ X lon nhất và tọa độ Y lớn nhất
                            }
                            else
                            {
                                pl.Add(new t3d.Point(partEdge.PointXmaxYmin));//add thêm Điểm có tọa độ X lon nhất và tọa độ Y nhỏ nhất
                                pl.Add(new t3d.Point(partEdge.PointXmaxYmax));//add thêm Điểm có tọa độ X lon nhất và tọa độ Y lớn nhất
                            }

                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                bolt_pl.Add(new t3d.Point(partEdge.PointXminYmin));
                                bolt_pl.Add(new t3d.Point(partEdge.PointXminYmax));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG chọn  Bolt Only
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pl, dimAttribute, dimSpace); // tạo DIM part phương Y
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương Y
                            }
                            else
                            {
                                bolt_pl.Add(new t3d.Point(partEdge.PointXmaxYmin));
                                bolt_pl.Add(new t3d.Point(partEdge.PointXmaxYmax));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pl, dimAttribute, dimSpace); // tạo DIM part phương Y
                                createDimension.CreateStraightDimensionSet_Y(viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương Y
                            }

                        } // Tấm đứng (front view), profile Plate
                        else if (Math.Abs(maxy.Y - miny.Y - chieuDayPlate) < 3 && partProfileType == "B")// Tấm ngang (front vỉew)
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();//Ds 2 điểm thấp nhất bên dưới
                            if (Control.ModifierKeys == Keys.Shift) // add điểm cho đim hướng xuống dưới
                            {
                                pl.Add(new t3d.Point(partEdge.PointXminYmin)); //add thêm Điểm có tọa độ X nhỏ nhất và tọa độ Y nhỏ nhất
                                pl.Add(new t3d.Point(partEdge.PointXmaxYmin)); //add thêm Điểm có tọa độ X lớn nhất và tọa độ Y nhỏ nhất
                            }
                            else
                            {
                                pl.Add(new t3d.Point(partEdge.PointXminYmax)); //add thêm Điểm có tọa độ X nhỏ nhất và tọa độ Y nhỏ nhất
                                pl.Add(new t3d.Point(partEdge.PointXmaxYmax)); //add thêm Điểm có tọa độ X lớn nhất và tọa độ Y nhỏ nhất
                            }
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                bolt_pl.Add(new t3d.Point(partEdge.PointXminYmin));
                                bolt_pl.Add(new t3d.Point(partEdge.PointXmaxYmin));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pl, dimAttribute, dimSpace); // tạo DIM part phương X
                                createDimension.CreateStraightDimensionSet_X_(viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương X
                            }
                            else
                            {
                                bolt_pl.Add(new t3d.Point(partEdge.PointXminYmax));
                                bolt_pl.Add(new t3d.Point(partEdge.PointXmaxYmax));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pl, dimAttribute, dimSpace); // tạo DIM part phương X
                                createDimension.CreateStraightDimensionSet_X(viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương X
                            }

                        } // Tấm ngang (front vỉew)
                        else if (Math.Abs(t3d.Distance.PointToPoint(minx, miny) - chieuDayPlate) < 3 && partProfileType == "B")// Tấm nghiêng phải
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                pl.Add(minx);//Điểm có tọa độ Y nhỏ nhất
                                pl.Add(maxy);//Điểm có tọa độ X lớn nhất
                            }
                            else
                            {
                                pl.Add(miny);//Điểm có tọa độ Y nhỏ nhất
                                pl.Add(maxx);//Điểm có tọa độ X lớn nhất
                            }

                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                bolt_pl.Add(minx);
                                bolt_pl.Add(maxy);
                                t3d.Vector vx = new Vector(maxx.X - miny.X, maxx.Y - miny.Y, 0);
                                t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_FX_(miny, maxx, viewBase, pl, dimAttribute, dimSpace); // tạo DIM part phương F
                                createDimension.CreateStraightDimensionSet_FX_(miny, maxx, viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương F
                            }
                            else
                            {
                                bolt_pl.Add(miny);
                                bolt_pl.Add(maxx);
                                t3d.Vector vx = new Vector(maxx.X - miny.X, maxx.Y - miny.Y, 0);
                                t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_FX(miny, maxx, viewBase, pl, dimAttribute, dimSpace); // tạo DIM part phương F
                                createDimension.CreateStraightDimensionSet_FX(miny, maxx, viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương F
                            }

                        } // Tấm nghiêng phải
                        else if (Math.Abs(t3d.Distance.PointToPoint(miny, maxx) - chieuDayPlate) < 3 && partProfileType == "B")// Tấm nghiêng trái
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                pl.Add(miny);//Điểm có tọa độ Y nhỏ nhất
                                pl.Add(minx);//Điểm có tọa độ X nhỏ nhất
                            }
                            else
                            {
                                pl.Add(maxy);//Điểm có tọa độ Y lớn nhất
                                pl.Add(maxx);//Điểm có tọa độ X lớn nhất
                            }

                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                bolt_pl.Add(miny);
                                bolt_pl.Add(minx);
                                t3d.Vector vx = new Vector(miny.X - minx.X, miny.Y - minx.Y, 0);
                                t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_FX(miny, minx, viewBase, pl, dimAttribute, dimSpace);//Tạo đường dim tổng
                                createDimension.CreateStraightDimensionSet_FX(miny, minx, viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương . Dùng CreateStraightDimensionSet_FX_ nếu muốn đặt dim ở hướng ngược lại
                            }
                            else
                            {
                                bolt_pl.Add(maxy);
                                bolt_pl.Add(maxx);
                                t3d.Vector vx = new Vector(miny.X - minx.X, miny.Y - minx.Y, 0);
                                t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    createDimension.CreateStraightDimensionSet_FX_(miny, minx, viewBase, pl, dimAttribute, dimSpace);//Tạo đường dim tổng
                                createDimension.CreateStraightDimensionSet_FX_(miny, minx, viewBase, bolt_pl, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương . Dùng CreateStraightDimensionSet_FX_ nếu muốn đặt dim ở hướng ngược lại
                            }

                        } // Tấm nghiêng trái

                        else if (partProfileType == "I" || partProfileType == "L" || partProfileType == "U" || partProfileType == "T" || partProfileType == "C" || partProfileType == "Z" || partProfileType == "RO" || partProfileType == "RU" || partProfileType == "M")
                        {
                            tsd.PointList bolt_pl = new PointList();
                            ArrayList arrayListCenterLine = mPart.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            //MessageBox.Show(pointCenterLineXmin.ToString() + "   " + pointCenterLineXmax.ToString());
                            List<t3d.Point> pointListPart = partEdge.List_Edge;
                            t3d.Point mid_YofPart = partEdge.PointMidLeft;
                            t3d.Point pointXmax0 = partEdge.PointXmax0;
                            t3d.Point pointXmin0 = partEdge.PointXmin0;
                            t3d.Point pointYmax0 = partEdge.PointYmax0;
                            t3d.Point pointYmin0 = partEdge.PointYmin0;
                            t3d.Point pointXminYmin = partEdge.PointXminYmin;
                            t3d.Point pointXminYmax = partEdge.PointXminYmax;
                            t3d.Point pointXmaxYmax = partEdge.PointXmaxYmax;
                            t3d.Point pointXmaxYmin = partEdge.PointXmaxYmin;

                            t3d.Point pointXmin1 = partEdge.PointXmin1;
                            t3d.Point pointXmax1 = partEdge.PointXmax1;
                            t3d.Point pointYmin1 = partEdge.PointYmin1;
                            t3d.Point pointYmax1 = partEdge.PointYmax1;
                            tsd.PointList dimOverall_X = new PointList(); //List point để dim part phương X
                            tsd.PointList dimOverall_Y = new PointList();//List point để dim part phương Y
                            tsd.PointList dim_Bolt_X = new PointList();//List point để dim Bolt phương X
                            tsd.PointList dim_Bolt_X_left = new PointList(); //List point để dim Bolt X (nằm ngang) phương Y ngược
                            tsd.PointList dim_Bolt_X_right = new PointList();//List point để dim Bolt X (nằm ngang) phương Y
                            tsd.PointList dim_Bolt_Y = new PointList();//List point để dim Bolt phương Y
                            tsd.PointList dim_Bolt_Y_bot = new PointList(); //List point để dim BoltY (nằm dọc) phương X ngược
                            tsd.PointList dim_Bolt_Y_top = new PointList(); //List point để dim BoltY (nằm dọc) phương X
                            tsd.PointList listpointBoltSkew = new PointList();//List point để dim Bolt xiêng

                            if (pointCenterLineXmin.Y < pointCenterLineXmax.Y && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) >
                                2 && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part nằm nghiêng Phải
                            {
                                dimOverall_X.Add(pointCenterLineXmin);
                                dimOverall_X.Add(pointCenterLineXmax);
                                dimOverall_Y.Add(pointYmax0);
                                dimOverall_Y.Add(pointXmax0);
                                dim_Bolt_X.Add(pointCenterLineXmin);
                                dim_Bolt_X.Add(pointCenterLineXmax);
                                dim_Bolt_Y_bot.Add(pointCenterLineXmin);
                                dim_Bolt_Y_bot.Add(pointCenterLineXmax);

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewBase, dimOverall_X, dimAttribute, dimSpace); // tạo DIM part phương FX
                                    createDimension.CreateStraightDimensionSet_FY_(pointCenterLineXmin, pointCenterLineXmax, viewBase, dimOverall_Y, dimAttribute, dimSpace); // tạo DIM part phương FY
                                }

                                foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                {
                                    try // đưa dim bolt vào try để trường hợp bolt chéo có 1 hàng sẽ vẫn tiếp tục chạy
                                    {
                                        Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                        List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY;
                                        List<t3d.Point> pointListBolt_Skew = tinh_Toan_Bolt.PointListBoltSkew;
                                        if (pointListBolt_XY != null)
                                        {
                                            foreach (t3d.Point p in pointListBolt_XY)
                                            {
                                                dim_Bolt_X.Add(p);
                                                dim_Bolt_Y_bot.Add(p);
                                            }
                                        }
                                        if (pointListBolt_Skew != null)
                                        {
                                            foreach (t3d.Point p in pointListBolt_Skew)
                                            {
                                                listpointBoltSkew.Add(p);
                                            }
                                            createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewBase, listpointBoltSkew, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FX
                                            listpointBoltSkew.Clear();
                                        }
                                    }
                                    catch
                                    { }
                                }

                                createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FX
                                createDimension.CreateStraightDimensionSet_FY_(pointCenterLineXmin, pointCenterLineXmax, viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FY
                            } //Part nằm nghiêng Phải
                            else if (pointCenterLineXmin.Y > pointCenterLineXmax.Y && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 2
                                && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part nằm nghiêng Trái
                            {
                                dimOverall_X.Add(pointCenterLineXmin);
                                dimOverall_X.Add(pointCenterLineXmax);
                                dimOverall_Y.Add(pointYmax0);
                                dimOverall_Y.Add(pointXmin0);
                                dim_Bolt_X.Add(pointCenterLineXmin);
                                dim_Bolt_X.Add(pointCenterLineXmax);
                                dim_Bolt_Y_bot.Add(pointCenterLineXmin);
                                dim_Bolt_Y_bot.Add(pointCenterLineXmax);

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewBase, dimOverall_X, dimAttribute, dimSpace); // tạo DIM part phương FX
                                    createDimension.CreateStraightDimensionSet_FY(pointCenterLineXmin, pointCenterLineXmax, viewBase, dimOverall_Y, dimAttribute, dimSpace); // tạo DIM part phương FY
                                }
                                foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                {
                                    try // đưa dim bolt vào try để trường hợp bolt chéo có 1 hàng sẽ vẫn tiếp tục chạy
                                    {
                                        Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                        List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY;
                                        List<t3d.Point> pointListBolt_Skew = tinh_Toan_Bolt.PointListBoltSkew;
                                        if (pointListBolt_XY != null)
                                        {
                                            foreach (t3d.Point p in pointListBolt_XY)
                                            {
                                                dim_Bolt_X.Add(p);
                                                dim_Bolt_Y_bot.Add(p);
                                            }
                                        }
                                        if (pointListBolt_Skew != null)
                                        {
                                            foreach (t3d.Point p in pointListBolt_Skew)
                                            {
                                                listpointBoltSkew.Add(p);
                                            }
                                            createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewBase, listpointBoltSkew, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FX
                                            listpointBoltSkew.Clear();
                                        }
                                    }
                                    catch
                                    { }
                                }
                                createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FX
                                createDimension.CreateStraightDimensionSet_FY(pointCenterLineXmin, pointCenterLineXmax, viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FY
                            } //Part nằm nghiêng Trái

                            else if (Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) < 2 && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 2) //Part nằm NGANG
                            {
                                double distanceBolt_PartXmin = 0;
                                double distanceBolt_PartXmax = 0;
                                PointList listpointBoltXYdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBoltXYdimY_ = new PointList() { pointXminYmax, pointXminYmin };
                                PointList listpointBoltXYdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };

                                PointList listpointBolt_YdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBolt_YdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };

                                PointList listpointBolt_XdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };
                                PointList listpointBolt_XdimY_ = new PointList() { pointXminYmax, pointXminYmin };

                                dimOverall_X.Add(pointXminYmax);
                                dimOverall_X.Add(pointXmaxYmax);
                                dimOverall_Y.Add(pointXminYmax);
                                dimOverall_Y.Add(pointXminYmin);

                                dim_Bolt_X.Add(pointXminYmax);
                                dim_Bolt_X.Add(pointXmaxYmax);
                                dim_Bolt_Y_bot.Add(pointXminYmax);
                                dim_Bolt_Y_bot.Add(pointXminYmin);
                                dim_Bolt_Y_top.Add(pointXmaxYmax);
                                dim_Bolt_Y_top.Add(pointXmaxYmin);
                                tsm.BoltGroup boltGroupXY = null;
                                tsm.BoltGroup boltGroupY = null;
                                tsm.BoltGroup boltGroupX = null;

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    if (Control.ModifierKeys != Keys.Alt)
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace); // tạo DIM part phương Y
                                }
                                List<BoltGroup> listBoltGroupXY = new List<BoltGroup>(); // list chứa các boltgroup Tròn
                                List<BoltGroup> listBoltGroupY = new List<BoltGroup>();// list chứa các boltgroup Nằm Dọc
                                List<BoltGroup> listBoltGroupX = new List<BoltGroup>();// list chứa các boltgroup Nằm Ngang

                                foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    boltGroupXY = tinh_Toan_Bolt.Bolt_XY; // 1 boltgroup Tròn
                                    boltGroupY = tinh_Toan_Bolt.Bolt_Y;// 1 boltgroup Nằm Dọc
                                    boltGroupX = tinh_Toan_Bolt.Bolt_X; // 1 boltgroup Nằm Ngang

                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY; // ds tất cả các điểm của boltgroup Tròn
                                    List<t3d.Point> pointListBolt_Y = tinh_Toan_Bolt.PointListBolt_Y; // ds tất cả các điểm của boltgroup Nằm Dọc
                                    List<t3d.Point> pointListBolt_X = tinh_Toan_Bolt.PointListBolt_X; // ds tất cả các điểm của boltgroup Ngang

                                    if (boltGroupXY != null) //nếu bolt Tròn không rỗng, xét TH có 1, 2, n bolt group
                                    {
                                        listBoltGroupXY.Add(boltGroupXY);
                                        foreach (t3d.Point p in pointListBolt_XY)
                                        {
                                            distanceBolt_PartXmin = Distance.PointToPoint(p, pointXminYmax); // thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim Y qua trái và phải là xong.
                                            distanceBolt_PartXmax = Distance.PointToPoint(p, pointXmaxYmax);// thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim Y qua trái và phải là xong.
                                            listpointBoltXYdimX.Add(p);
                                            listpointBoltXYdimY_.Add(p);
                                            listpointBoltXYdimY.Add(p);
                                        }
                                    }

                                    else if (boltGroupY != null)
                                    {
                                        listBoltGroupY.Add(boltGroupY);
                                        foreach (t3d.Point p in pointListBolt_Y)
                                        {
                                            if (p.Y > pointCenterLineXmin.Y) //Nếu nằm TRÊN đường centerline
                                            {
                                                listpointBolt_YdimX.Add(p);
                                            }
                                            else //Nếu nằm DƯỚI đường centerline
                                            {
                                                listpointBolt_YdimX_.Add(p);
                                            }
                                        }
                                    }

                                    else if (pointListBolt_X != null)
                                    {
                                        MessageBox.Show("Có bolt nằm ngang");
                                        break;
                                    }
                                }

                                //DIM cho bolt TRÒN
                                createDimension.CreateStraightDimensionSet_X(viewBase, listpointBoltXYdimX, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt TRÒN
                                if (listBoltGroupXY.Count == 1) // nếu chỉ có 1 BoltGroup TRÒN
                                {
                                    if (distanceBolt_PartXmin < distanceBolt_PartXmax) //Nếu kc từ bolt đến mép trái part nhỏ hơn, thì tạo dim qua trái
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBoltXYdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua trái
                                    else
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, listpointBoltXYdimY, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua phải
                                }
                                else if (listBoltGroupXY.Count == 2) // nếu có 2 BoltGroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBoltXYdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua trái
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, listpointBoltXYdimY, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua phải
                                }
                                else //nếu có nhiều boltgroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBoltXYdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim);// tạo DIM Bolt phương Y qua trái
                                }
                                //DIM cho bolt NẰM DỌC
                                if (boltGroupY != null)
                                {
                                    if (listpointBolt_YdimX.Count > 2)
                                        createDimension.CreateStraightDimensionSet_X(viewBase, listpointBolt_YdimX, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                    if (listpointBolt_YdimX_.Count > 2)
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, listpointBolt_YdimX_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                }

                            } //Part nằm NGANG
                            else if (Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) < 2 && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part ĐỨNG
                            {
                                double distanceBolt_PartYmin = 0;
                                double distanceBolt_PartYmax = 0;
                                PointList listpointBoltXYdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBoltXYdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };
                                PointList listpointBoltXYdimY_ = new PointList() { pointXminYmax, pointXminYmin };
                                PointList listpointBoltXYdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };

                                PointList listpointBolt_YdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBolt_YdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };

                                PointList listpointBolt_XdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };
                                PointList listpointBolt_XdimY_ = new PointList() { pointXminYmax, pointXminYmin };

                                dimOverall_X.Add(pointXminYmax);
                                dimOverall_X.Add(pointXmaxYmax);
                                dimOverall_Y.Add(pointXminYmax);
                                dimOverall_Y.Add(pointXminYmin);

                                dim_Bolt_X.Add(pointXminYmax);
                                dim_Bolt_X.Add(pointXmaxYmax);
                                dim_Bolt_Y_bot.Add(pointXminYmax);
                                dim_Bolt_Y_bot.Add(pointXminYmin);
                                dim_Bolt_Y_top.Add(pointXmaxYmax);
                                dim_Bolt_Y_top.Add(pointXmaxYmin);
                                tsm.BoltGroup boltGroupXY = null;
                                tsm.BoltGroup boltGroupY = null;
                                tsm.BoltGroup boltGroupX = null;
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace); // tạo DIM part phương Y
                                }
                                List<BoltGroup> listBoltGroupXY = new List<BoltGroup>(); // list chứa các boltgroup Tròn
                                List<BoltGroup> listBoltGroupY = new List<BoltGroup>();// list chứa các boltgroup Nằm Dọc
                                List<BoltGroup> listBoltGroupX = new List<BoltGroup>();// list chứa các boltgroup Nằm Ngang

                                foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    boltGroupXY = tinh_Toan_Bolt.Bolt_XY; // 1 boltgroup Tròn
                                    boltGroupY = tinh_Toan_Bolt.Bolt_Y;// 1 boltgroup Nằm Dọc
                                    boltGroupX = tinh_Toan_Bolt.Bolt_X; // 1 boltgroup Nằm Ngang

                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY; // ds tất cả các điểm của boltgroup Tròn
                                    List<t3d.Point> pointListBolt_Y = tinh_Toan_Bolt.PointListBolt_Y; // ds tất cả các điểm của boltgroup Nằm Dọc
                                    List<t3d.Point> pointListBolt_X = tinh_Toan_Bolt.PointListBolt_X; // ds tất cả các điểm của boltgroup Ngang

                                    if (boltGroupXY != null) //nếu bolt Tròn không rỗng, xét TH có 1, 2, n bolt group
                                    {
                                        listBoltGroupXY.Add(boltGroupXY);
                                        foreach (t3d.Point p in pointListBolt_XY)
                                        {
                                            distanceBolt_PartYmin = Distance.PointToPoint(p, pointXminYmin); // thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim X lên trên xuống dưới.
                                            distanceBolt_PartYmax = Distance.PointToPoint(p, pointXminYmax);// thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim X lên trên xuống dưới.
                                            listpointBoltXYdimX.Add(p);
                                            listpointBoltXYdimX_.Add(p);
                                            listpointBoltXYdimY_.Add(p);
                                        }
                                    }

                                    else if (boltGroupX != null)
                                    {
                                        listBoltGroupY.Add(boltGroupY);
                                        foreach (t3d.Point p in pointListBolt_X)
                                        {
                                            if (p.X < pointCenterLineXmin.X) //Nếu nằm BÊN TRÁI đường centerline
                                            {
                                                listpointBolt_XdimY_.Add(p);
                                            }
                                            else //Nếu nằm BÊN PHẢI đường centerline
                                            {
                                                listpointBolt_XdimY.Add(p);
                                            }
                                        }
                                    }

                                    else if (boltGroupY != null)
                                    {
                                        MessageBox.Show("Có bolt nằm doc");
                                    }
                                }

                                //DIM cho bolt TRÒN
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBoltXYdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt TRÒN
                                if (listBoltGroupXY.Count == 1) // nếu chỉ có 1 BoltGroup TRÒN
                                {
                                    if (distanceBolt_PartYmin > distanceBolt_PartYmax) //Nếu kc từ bolt đến mép DƯỚI part lớn hơn, thì tạo dim lên TRÊN
                                        createDimension.CreateStraightDimensionSet_X(viewBase, listpointBoltXYdimX, dimAttribute, dimSpace - dimSpaceDim2Dim);   // tạo DIM Bolt phương X lên trên
                                    else
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, listpointBoltXYdimX_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X xuống dưới
                                }
                                else if (listBoltGroupXY.Count == 2) // nếu có 2 BoltGroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewBase, listpointBoltXYdimX, dimAttribute, dimSpace - dimSpaceDim2Dim); //   tạo DIM Bolt phương X qua lên trên
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, listpointBoltXYdimX_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X qua xuống dưới
                                }
                                else //nếu có nhiều boltgroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewBase, listpointBoltXYdimX, dimAttribute, dimSpace - dimSpaceDim2Dim);// tạo DIM Bolt phương Y qua trái
                                }
                                //DIM cho bolt NẰM NGANG
                                if (boltGroupX != null)
                                {
                                    if (listpointBolt_XdimY_.Count > 2)
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBolt_XdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua trái cho bolt NẰM NGANG
                                    if (listpointBolt_XdimY.Count > 2)
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, listpointBolt_XdimY, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua phải cho bolt NẰM NGANG
                                }


                            } //Part ĐỨNG
                            else if (Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) < 2 && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) < 2) //SECTION PROFILE
                            {
                                double distanceBolt_PartXmin = 0;
                                double distanceBolt_PartXmax = 0;
                                PointList listpointBoltXYdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBoltXYdimY_ = new PointList() { pointXminYmax, pointXminYmin };
                                PointList listpointBoltXYdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };

                                PointList listpointBolt_YdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBolt_YdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };

                                PointList listpointBolt_XdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };
                                PointList listpointBolt_XdimY_ = new PointList() { pointXminYmax, pointXminYmin };

                                //dimOverall_X.Add(pointXminYmax);
                                //dimOverall_X.Add(pointXmaxYmax);
                                //dimOverall_Y.Add(pointXminYmax);
                                //dimOverall_Y.Add(pointXminYmin);

                                //dim_Bolt_X.Add(pointXminYmax);
                                //dim_Bolt_X.Add(pointXmaxYmax);
                                //dim_Bolt_Y_bot.Add(pointXminYmax);
                                //dim_Bolt_Y_bot.Add(pointXminYmin);
                                //dim_Bolt_Y_top.Add(pointXmaxYmax);
                                //dim_Bolt_Y_top.Add(pointXmaxYmin);
                                tsm.BoltGroup boltGroupXY = null;
                                tsm.BoltGroup boltGroupY = null;
                                tsm.BoltGroup boltGroupX = null;
                                List<BoltGroup> listBoltGroupXY = new List<BoltGroup>(); // list chứa các boltgroup Tròn
                                List<BoltGroup> listBoltGroupY = new List<BoltGroup>();// list chứa các boltgroup Nằm Dọc
                                List<BoltGroup> listBoltGroupX = new List<BoltGroup>();// list chứa các boltgroup Nằm Ngang

                                foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    //boltGroupXY = tinh_Toan_Bolt.Bolt_XY; // 1 boltgroup Tròn  // section profile thì không thấy bolt tròn
                                    boltGroupY = tinh_Toan_Bolt.Bolt_Y;// 1 boltgroup Nằm Dọc
                                    boltGroupX = tinh_Toan_Bolt.Bolt_X; // 1 boltgroup Nằm Ngang

                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY; // ds tất cả các điểm của boltgroup Tròn
                                    List<t3d.Point> pointListBolt_Y = tinh_Toan_Bolt.PointListBolt_Y; // ds tất cả các điểm của boltgroup Nằm Dọc
                                    List<t3d.Point> pointListBolt_X = tinh_Toan_Bolt.PointListBolt_X; // ds tất cả các điểm của boltgroup Ngang
                                    if (boltGroupY != null)
                                    {
                                        listBoltGroupY.Add(boltGroupY);
                                        foreach (t3d.Point p in pointListBolt_Y)
                                        {
                                            if (p.Y > pointCenterLineXmin.Y) //Nếu nằm TRÊN đường centerline
                                            {
                                                listpointBolt_YdimX.Add(p);
                                            }
                                            else //Nếu nằm DƯỚI đường centerline
                                            {
                                                listpointBolt_YdimX_.Add(p);
                                            }
                                        }
                                    }
                                    else if (pointListBolt_X != null)
                                    {
                                        listBoltGroupX.Add(boltGroupX);
                                        foreach (t3d.Point p in pointListBolt_X)
                                        {
                                            if (p.X > pointCenterLineXmin.X) //Nếu nằm TRÊN đường centerline
                                            {
                                                listpointBolt_XdimY_.Add(p);
                                            }
                                            else //Nếu nằm DƯỚI đường centerline
                                            {
                                                listpointBolt_XdimY_.Add(p);
                                            }
                                        }
                                    }
                                }
                                //Xét SECTION của Profile U (C) hay L(V) hay I mà tạo dim cho phù hợp:
                                if (partProfileType == "U")
                                {
                                    //tạo ds các điểm có cùng tọa độ X min, Nếu số phần từ là 4 thì C quay lưng về bên trái và ngược lại.
                                    List<t3d.Point> pointListPart_minX = pointListPart.FindAll(p => Math.Abs(p.X - pointXmin0.X) < 0.1).ToList();
                                    dimOverall_X.Add(pointYmax0);
                                    dimOverall_X.Add(pointYmax1);
                                    //Nếu C quay lưng qua trái
                                    if (pointListPart_minX.Count == 4)
                                    {
                                        dimOverall_Y.Add(pointYmax1);
                                        dimOverall_Y.Add(pointYmin1);
                                    }
                                    //Nếu C quay lưng qua phải
                                    else
                                    {
                                        dimOverall_Y.Add(pointYmax0);
                                        dimOverall_Y.Add(pointYmin0);
                                    }
                                    dim_Bolt_Y_top.Add(pointYmax0);
                                    dim_Bolt_Y_top.Add(pointYmax1);
                                    dim_Bolt_Y_bot.Add(pointYmin0);
                                    dim_Bolt_Y_bot.Add(pointYmin1);
                                    //Nếu C quay lưng qua trái
                                    if (pointListPart_minX.Count == 4)
                                    {
                                        dim_Bolt_X.Add(pointYmax1);
                                        dim_Bolt_X.Add(pointYmin1);
                                    }
                                    //Nếu C quay lưng qua phải
                                    else
                                    {
                                        dim_Bolt_X.Add(pointYmax0);
                                        dim_Bolt_X.Add(pointYmin0);
                                    }
                                    //tạo dim X lên trên
                                    createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace);
                                    //tạo dim Y qua trái
                                    if (pointListPart_minX.Count == 4)
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                    else //tạo dim Y qua phải
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                    foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                    {
                                        Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                        //Tìm bolt theo phương XY
                                        tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                        tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                        //wph.SetCurrentTransformationPlane(viewplane);
                                        if (bolt_x != null)
                                        {
                                            bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_x.BoltPositions)
                                                dim_Bolt_X.Add(p);
                                        }
                                        if (bolt_y != null)
                                        {
                                            bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_y.BoltPositions)
                                            {
                                                if (p.Y > mid_YofPart.Y)
                                                    dim_Bolt_Y_top.Add(p);
                                                else
                                                    dim_Bolt_Y_bot.Add(p);
                                            }
                                        }
                                    }
                                    if (dim_Bolt_Y_top.Count != 2)
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - dimSpaceDim2Dim);

                                    if (dim_Bolt_Y_bot.Count != 2)
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim);

                                    if (dim_Bolt_X.Count != 2)
                                    {
                                        if (pointListPart_minX.Count == 4)
                                            createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                        else //tạo dim Y qua phải
                                            createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                    }
                                }
                                else if (partProfileType == "L")
                                {
                                    //tạo ds các điểm có cùng tọa độ X min, Nếu số phần từ là 4 thì C quay lưng về bên trái và ngược lại.
                                    List<t3d.Point> pointListPart_minX = pointListPart.FindAll(p => Math.Abs(p.X - pointXmin0.X) < 0.1).ToList();
                                    List<t3d.Point> listCheckXmaxYmax = pointListPart.FindAll(p => p == pointXmaxYmax).ToList();
                                    List<t3d.Point> listCheckXminYmax = pointListPart.FindAll(p => p == pointXminYmax).ToList();
                                    List<t3d.Point> listCheckXmaxYmin = pointListPart.FindAll(p => p == pointXmaxYmin).ToList();
                                    List<t3d.Point> listCheckXminYmin = pointListPart.FindAll(p => p == pointXminYmin).ToList();

                                    //Nếu L không có point XmaxYmax
                                    if (listCheckXmaxYmax.Count == 0)
                                    {
                                        //MessageBox.Show("XmaxYmax");
                                        dimOverall_X.Add(pointXmaxYmin);
                                        dimOverall_X.Add(pointXminYmin);
                                        dimOverall_Y.Add(pointXminYmax);
                                        dimOverall_Y.Add(pointXminYmin);
                                        dim_Bolt_X.Add(pointXminYmax);
                                        dim_Bolt_X.Add(pointXminYmin);
                                        dim_Bolt_Y.Add(pointXmaxYmin);
                                        dim_Bolt_Y.Add(pointXminYmin);
                                        //tạo dim X huong xuong
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, dimOverall_X, dimAttribute, dimSpace);
                                        //tạo dim Y qua trai
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                    }
                                    //Nếu C quay lưng qua phải
                                    else if (listCheckXminYmax.Count == 0)
                                    {
                                        //MessageBox.Show("XminYmax");
                                        dimOverall_X.Add(pointXmaxYmin);
                                        dimOverall_X.Add(pointXminYmin);
                                        dimOverall_Y.Add(pointXmaxYmax);
                                        dimOverall_Y.Add(pointXmaxYmin);
                                        //dim_Bolt_X.Add(pointXmaxYmax);
                                        dim_Bolt_X.Add(pointXmaxYmin);
                                        dim_Bolt_Y.Add(pointXmaxYmin);
                                        dim_Bolt_Y.Add(pointXminYmin);
                                        //tạo dim X huong xuong
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, dimOverall_X, dimAttribute, dimSpace);
                                        //tạo dim Y qua phai
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                    }
                                    else if (listCheckXmaxYmin.Count == 0)
                                    {
                                        //MessageBox.Show("XmaxYmin");
                                        dimOverall_X.Add(pointXmaxYmax);
                                        dimOverall_X.Add(pointXminYmax);
                                        dimOverall_Y.Add(pointXminYmax);
                                        dimOverall_Y.Add(pointXminYmin);
                                        dim_Bolt_X.Add(pointXminYmax);
                                        dim_Bolt_X.Add(pointXminYmin);
                                        dim_Bolt_Y.Add(pointXmaxYmax);
                                        dim_Bolt_Y.Add(pointXminYmax);
                                        //tạo dim X huong len
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace);
                                        //tạo dim Y qua trai
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                    }
                                    else if (listCheckXminYmin.Count == 0)
                                    {
                                        //MessageBox.Show("XminYmin");
                                        dimOverall_X.Add(pointXmaxYmax);
                                        dimOverall_X.Add(pointXminYmax);
                                        dimOverall_Y.Add(pointXmaxYmax);
                                        dimOverall_Y.Add(pointXmaxYmin);
                                        dim_Bolt_X.Add(pointXmaxYmax);
                                        dim_Bolt_X.Add(pointXmaxYmin);
                                        dim_Bolt_Y.Add(pointXmaxYmax);
                                        dim_Bolt_Y.Add(pointXminYmax);
                                        //tạo dim X huong len
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace);
                                        //tạo dim Y qua phai
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                    }

                                    foreach (tsm.BoltGroup bgr in mPart.GetBolts())
                                    {
                                        Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                        //Tìm bolt theo phương XY
                                        tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                        tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                        //wph.SetCurrentTransformationPlane(viewplane);
                                        if (bolt_x != null)
                                        {
                                            bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_x.BoltPositions)
                                                dim_Bolt_X.Add(p);
                                        }
                                        if (bolt_y != null)
                                        {
                                            bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_y.BoltPositions)
                                            {
                                                dim_Bolt_Y.Add(p);
                                            }
                                        }
                                    }
                                    if (listCheckXmaxYmax.Count == 0)
                                    {
                                        //tạo dim Bolt X huong xuong
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                        //tạo dim Bolt Y qua trai
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                    }
                                    else if (listCheckXminYmax.Count == 0)
                                    {
                                        //tạo dim Bolt X huong xuong
                                        createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                        //tạo dim Bolt Y qua phai
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                    }
                                    else if (listCheckXmaxYmin.Count == 0)
                                    {
                                        //tạo dim Bolt X huong len
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                        //tạo dim Bolt Y qua trai
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                    }
                                    else if (listCheckXminYmin.Count == 0)
                                    {
                                        //tạo dim Bolt X huong len
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                        //tạo dim Bolt Y qua phai
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                    }
                                }
                                else
                                {
                                    //Dim total X Y cho part
                                    if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    {
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace); // tạo DIM part phương X
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace); // tạo DIM part phương Y
                                    }
                                    //DIM cho bolt NẰM DỌC
                                    if (boltGroupY != null)
                                    {
                                        if (listpointBolt_YdimX.Count > 2)
                                            createDimension.CreateStraightDimensionSet_X(viewBase, listpointBolt_YdimX, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                        if (listpointBolt_YdimX_.Count > 2)
                                            createDimension.CreateStraightDimensionSet_X_(viewBase, listpointBolt_YdimX_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                    }
                                    //DIM cho bolt NẰM NGANG
                                    if (boltGroupX != null)
                                    {
                                        if (listpointBolt_XdimY.Count > 2)
                                            createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBolt_XdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                        if (listpointBolt_XdimY_.Count > 2)
                                            createDimension.CreateStraightDimensionSet_Y_(viewBase, listpointBolt_XdimY_, dimAttribute, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                    }
                                }

                            } //Part SECTION
                            else //Tấm thấy bolt tròn
                            {
                                tsd.PointList dimXPart = new PointList();
                                tsd.PointList dimYPart = new PointList();
                                tsd.PointList Boltlist_Bolt_XY = new PointList();
                                tsd.PointList Boltlist_Bolt_Y = new PointList();
                                tsd.PointList Boltlist_Bolt_X = new PointList();
                                tsd.PointList Boltlist_Bolt_skew = new PointList();
                                //kiểm tra xem đối tượng này có phải là part không ?
                                if (drobj is tsd.Part)
                                {
                                    tsm.ModelObjectEnumerator boltenum = mPart.GetBolts();
                                    foreach (tsm.BoltGroup bgr in boltenum)
                                    {
                                        Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                        //Tìm bolt theo phương XY
                                        tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY;
                                        tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                        wph.SetCurrentTransformationPlane(viewplane);
                                        if (bolt_xy != null)
                                        {
                                            bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_xy.BoltPositions)
                                            {
                                                Boltlist_Bolt_skew.Add(p);
                                                Boltlist_Bolt_XY.Add(p);
                                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, maxy.Y));
                                                Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, maxy.Y));
                                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));
                                            }
                                        }
                                        if (bolt_y != null)
                                        {
                                            bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_y.BoltPositions)
                                            {
                                                Boltlist_Bolt_Y.Add(p);
                                                Boltlist_Bolt_Y.Add(new t3d.Point(minx.X, maxy.Y));
                                                Boltlist_Bolt_Y.Add(new t3d.Point(maxx.X, maxy.Y));
                                            }
                                        }
                                    }

                                    dimXPart.Add(new t3d.Point(minx.X, maxy.Y));// góc trên bên trái
                                    dimXPart.Add(new t3d.Point(maxx.X, maxy.Y));// góc trên bên phải
                                    dimYPart.Add(new t3d.Point(minx.X, maxy.Y)); // góc trên bên trái
                                    dimYPart.Add(new t3d.Point(minx.X, miny.Y)); // góc dưới bên trái

                                    if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    {
                                        createDimension.CreateStraightDimensionSet_X(viewBase, dimXPart, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương X
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dimYPart, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương Y
                                    }
                                    createDimension.CreateStraightDimensionSet_X(viewBase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM BOlt phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương -Y
                                }
                            } //Tấm thấy bolt tròn
                            wph.SetCurrentTransformationPlane(new TransformationPlane());
                        }

                        else if (Math.Abs(maxx.X - minx.X - chieuDayPlate) > 3 && partProfileType == "B") // Tấm THẤY BOLT TRÒN, profile Plate
                        {
                            tsd.PointList bolt_pl = new PointList();
                            ArrayList arrayListCenterLine = mPart.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            //MessageBox.Show(pointCenterLineXmin.ToString() + "   " + pointCenterLineXmmax.ToString());
                            t3d.Point pointXmax = partEdge.PointXmax0;
                            t3d.Point pointXmin = partEdge.PointXmin0;
                            t3d.Point pointYmax = partEdge.PointYmax0;
                            t3d.Point pointYmin = partEdge.PointYmin0;
                            t3d.Point pointXminYmin = partEdge.PointXminYmin;
                            t3d.Point pointXminYmax = partEdge.PointXminYmax;
                            t3d.Point pointXmaxYmax = partEdge.PointXmaxYmax;
                            t3d.Point pointXmaxYmin = partEdge.PointXmaxYmin;
                            tsd.PointList listpointPartX = new PointList();
                            tsd.PointList listpointPartY_ = new PointList();
                            tsd.PointList listpointBoltX = new PointList();
                            tsd.PointList listpointBoltY_ = new PointList();
                            tsd.PointList listpointBoltY = new PointList();

                            tsd.PointList dimXPart = new PointList();
                            tsd.PointList dimYPart = new PointList();
                            tsd.PointList Boltlist_Bolt_XY = new PointList();
                            tsd.PointList Boltlist_Bolt_Y = new PointList();
                            tsd.PointList Boltlist_Bolt_X = new PointList();
                            tsd.PointList Boltlist_Bolt_skew = new PointList();
                            tsm.ModelObjectEnumerator boltenum = mPart.GetBolts();
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                if (Control.ModifierKeys == Keys.Shift)
                                {
                                    dimXPart.Add(new t3d.Point(minx.X, miny.Y));// góc duoi bên trái
                                    dimXPart.Add(new t3d.Point(maxx.X, miny.Y));// góc duoi bên phải
                                    dimYPart.Add(new t3d.Point(maxx.X, maxy.Y)); // góc trên bên phai
                                    dimYPart.Add(new t3d.Point(maxx.X, miny.Y)); // góc dưới bên phai
                                    Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, miny.Y));// góc dưới bên phải
                                    Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));// góc dưới bên trái
                                    Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, maxy.Y));// góc trên bên phải
                                }
                                else
                                {
                                    dimXPart.Add(new t3d.Point(minx.X, maxy.Y));// góc trên bên trái
                                    dimXPart.Add(new t3d.Point(maxx.X, maxy.Y));// góc trên bên phải
                                    dimYPart.Add(new t3d.Point(minx.X, maxy.Y)); // góc trên bên trái
                                    dimYPart.Add(new t3d.Point(minx.X, miny.Y)); // góc dưới bên trái
                                    Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, maxy.Y));// góc trên bên trái
                                    Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, maxy.Y));// góc trên bên phải
                                    Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));// góc dưới bên trái
                                }


                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY;
                                wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_xy != null)
                                {
                                    bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_xy.BoltPositions)
                                    {
                                        Boltlist_Bolt_XY.Add(p);
                                    }
                                }
                            }
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, dimXPart, cb_DimentionAttributes.Text, 2 * dimSpaceDim2Dim); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, dimYPart, cb_DimentionAttributes.Text, 2 * dimSpaceDim2Dim); // tạo DIM part phương Y
                                }
                                createDimension.CreateStraightDimensionSet_X_(viewBase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpaceDim2Dim); // tạo DIM BOlt phương X
                                createDimension.CreateStraightDimensionSet_Y(viewBase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpaceDim2Dim); // tạo DIM bolt phương -Y
                            }
                            else
                            {
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewBase, dimXPart, cb_DimentionAttributes.Text, 2 * dimSpaceDim2Dim); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, dimYPart, cb_DimentionAttributes.Text, 2 * dimSpaceDim2Dim); // tạo DIM part phương Y
                                }
                                createDimension.CreateStraightDimensionSet_X(viewBase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpaceDim2Dim); // tạo DIM BOlt phương X
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpaceDim2Dim); // tạo DIM bolt phương -Y  
                            }
                        }//Tấm thấy bolt tròn
                        wph.SetCurrentTransformationPlane(new TransformationPlane());
                    }
                    if (drobj is tsd.Bolt)
                    {
                        tsd.Bolt drBolt = drobj as tsd.Bolt;
                        tsd.ViewBase viewbase = drobj.GetView();
                        tsd.View viewcur = drobj.GetView() as tsd.View;
                        wph.SetCurrentTransformationPlane(new TransformationPlane());
                        tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        if (chbox_ClearCurrentDim.Checked)
                            clearDrawingObjects.ClearDim(viewcur);
                        //model.CommitChanges();
                        tsd.PointList bolt_pl = new PointList();
                        //tsm.ModelObject myPart = new tsm.Model().SelectModelObject(modelObject.ModelIdentifier) as tsm.ModelObject;
                        tsm.BoltGroup boltGroup = new tsm.Model().SelectModelObject(drBolt.ModelIdentifier) as tsm.BoltGroup;
                        t3d.Point firstPoint = boltGroup.FirstPosition;
                        t3d.Point secondPoint = boltGroup.SecondPosition;
                        foreach (t3d.Point p in boltGroup.BoltPositions)
                        {
                            //MessageBox.Show(p.ToString());
                            bolt_pl.Add(p);
                        }
                        if (chbox_SkewBolts.Checked)
                        {
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                createDimension.CreateStraightDimensionSet_FX_(firstPoint, secondPoint, viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace);
                                createDimension.CreateStraightDimensionSet_FY(firstPoint, secondPoint, viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace);
                            }
                            else
                            {
                                createDimension.CreateStraightDimensionSet_FX(firstPoint, secondPoint, viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace);
                                createDimension.CreateStraightDimensionSet_FY_(firstPoint, secondPoint, viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace);
                            }
                        }
                        else
                        {
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                try
                                {
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương Y
                                    createDimension.CreateStraightDimensionSet_X_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương X
                                }
                                catch
                                { createDimension.CreateStraightDimensionSet_X_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); }
                            }
                            else
                            {
                                try
                                {
                                    createDimension.CreateStraightDimensionSet_Y(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương Y
                                    createDimension.CreateStraightDimensionSet_X(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương X
                                }
                                catch
                                { createDimension.CreateStraightDimensionSet_X(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); }
                            }
                        }
                    }
                }
                if (chbox_AddPartMark.Checked)
                    Add_Mark();
            }
            catch { }
        Ketthuc:;
            drawingHandler.GetActiveDrawing().CommitChanges();
        }

        public void DimGratingCheckerPlate()
        {
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            wph.SetCurrentTransformationPlane(new TransformationPlane());
            CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            tsdui.Picker c_picker = drawingHandler.GetPicker();
            int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
            int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
            PointList pointsDimTop = new PointList();
            PointList pointsDimBot = new PointList();
            PointList pointsDimLeft = new PointList();
            PointList pointsDimRight = new PointList();

            List<t3d.Point> pointsOnLineTop = new List<t3d.Point>();
            List<t3d.Point> pointsOnLineBot = new List<t3d.Point>();
            List<t3d.Point> pointsOnLineLeft = new List<t3d.Point>();
            List<t3d.Point> pointsOnLineRight = new List<t3d.Point>();
            List<t3d.Point> pointsInsight = new List<t3d.Point>(); //danh sách tất cả các điểm không nằm trên 4 cạnh line
            try
            {
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drpart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                        //Lấy view hiện hành thông qua drpart
                        tsd.View viewcur = drpart.GetView() as tsd.View;
                        tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.ViewBase viewbase = drpart.GetView();
                        double viewScale = viewcur.Attributes.Scale;
                        if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                        {
                            dimSpace = Convert.ToInt32(viewScale * 10);
                            dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                        }
                        //Tính điểm định nghĩa của part theo bản vẽ
                        Part_Edge part_edge = new Part_Edge(viewcur, mpart);
                        List<t3d.Point> pointsPartEdge = part_edge.List_Edge;
                        t3d.Point pointYmax = part_edge.PointYmax0;
                        t3d.Point pointYmin = part_edge.PointYmin0;
                        t3d.Point pointXmax = part_edge.PointXmax0;
                        t3d.Point pointXmin = part_edge.PointXmin0;

                        t3d.Line lineTop = new t3d.Line(pointYmax, new t3d.Point(pointYmax.X + 1000, pointYmax.Y));
                        t3d.Line lineBot = new t3d.Line(pointYmin, new t3d.Point(pointYmin.X + 1000, pointYmin.Y));
                        t3d.Line lineLeft = new t3d.Line(pointXmin, new t3d.Point(pointXmin.X, pointXmin.Y + 1000));
                        t3d.Line lineRight = new t3d.Line(pointXmax, new t3d.Point(pointXmax.X, pointYmax.Y + 1000));

                        foreach (t3d.Point point in pointsPartEdge)
                        {
                            if (Distance.PointToLine(point, lineTop) <= 3) //Điểm on line Top
                            {
                                pointsOnLineTop.Add(point);
                                pointsDimTop.Add(point);
                            }
                            if (Distance.PointToLine(point, lineBot) <= 3) //Điểm on line Bot
                            {
                                pointsOnLineBot.Add(point);
                                pointsDimBot.Add(point);
                            }
                            if (Distance.PointToLine(point, lineLeft) <= 3) //Điểm on line Left
                            {
                                pointsOnLineLeft.Add(point);
                                pointsDimLeft.Add(point);
                            }
                            if (Distance.PointToLine(point, lineRight) <= 3) //Điểm on line Right
                            {
                                pointsOnLineRight.Add(point);
                                pointsDimRight.Add(point);
                            }
                            else if (Distance.PointToLine(point, lineTop) > 3 && Distance.PointToLine(point, lineBot) > 3 && Distance.PointToLine(point, lineLeft) > 3 && Distance.PointToLine(point, lineRight) > 3)
                            {
                                pointsInsight.Add(point);
                            }
                        }
                        t3d.Point minXLineTop = pointsOnLineTop.OrderBy(point => point.X).ToList()[0];
                        t3d.Point maxXLineTop = pointsOnLineTop.OrderByDescending(point => point.X).ToList()[0];
                        t3d.Point minXLineBot = pointsOnLineBot.OrderBy(point => point.X).ToList()[0];
                        t3d.Point maxXLineBot = pointsOnLineBot.OrderByDescending(point => point.X).ToList()[0];
                        t3d.Point minYLineLeft = pointsOnLineLeft.OrderBy(point => point.Y).ToList()[0];
                        t3d.Point maxYLineLeft = pointsOnLineLeft.OrderByDescending(point => point.Y).ToList()[0];
                        t3d.Point minYLineRight = pointsOnLineRight.OrderBy(point => point.Y).ToList()[0];
                        t3d.Point maxYLineRight = pointsOnLineRight.OrderByDescending(point => point.Y).ToList()[0];

                        List<t3d.Point> pointsCornerLeftTop = new List<t3d.Point>(); //Tập hợn các điểm góc trên bên trái có X nhỏ hơn minXLineTop và có Y lớn hơn maxYLineLeft
                        List<t3d.Point> pointsCornerLeftTop_X = new List<t3d.Point>(); //Tập hợn các điểm góc trên bên trái không trùng tọa độ X
                        List<t3d.Point> pointsCornerLeftTop_Y = new List<t3d.Point>(); //Tập hợn các điểm góc trên bên trái không trùng tọa độ X
                        List<t3d.Point> pointsCornerRightTop = new List<t3d.Point>(); //Tập hợn các điểm góc trên bên phải có X lơn hơn maxXLineTop và có Y lớn hơn maxYLineRight
                        List<t3d.Point> pointsCornerRightTop_X = new List<t3d.Point>(); //Tập hợn các điểm góc trên bên trái không trùng tọa độ X
                        List<t3d.Point> pointsCornerRightTop_Y = new List<t3d.Point>(); //Tập hợn các điểm góc trên bên trái không trùng tọa độ X
                        List<t3d.Point> pointsCornerRightBot = new List<t3d.Point>(); //Tập hợn các điểm góc dưới bên Phải có X lớn hơn maxXLineBot và có Y nhỏ hơn minYLineRight
                        List<t3d.Point> pointsCornerRightBot_X = new List<t3d.Point>(); //Tập hợn các điểm góc dưới bên trái không trùng tọa độ X
                        List<t3d.Point> pointsCornerRightBot_Y = new List<t3d.Point>(); //Tập hợn các điểm góc dưới bên trái không trùng tọa độ Y
                        List<t3d.Point> pointsCornerLeftBot = new List<t3d.Point>(); //Tập hợn các điểm góc dưới bên trái có X nhỏ hơn minXLineBot và có Y nhỏ hơn minYLineLeft
                        List<t3d.Point> pointsCornerLeftBot_X = new List<t3d.Point>(); //Tập hợn các điểm góc dưới bên trái không trùng tọa độ X
                        List<t3d.Point> pointsCornerLeftBot_Y = new List<t3d.Point>(); //Tập hợn các điểm góc dưới bên trái không trùng tọa độ X
                        #region Phần này để thêm các điểm nằm trong ô góc trên bên trái (tạo bởi maxYLineLeft và minXLineTop)
                        if (Math.Abs(Distance.PointToLine(minXLineTop, lineLeft)) >= 3)
                        {
                            pointsDimTop.Add(maxYLineLeft); //Thêm điểm cao nhất bên trái vào điểm dim TOP
                            pointsDimLeft.Add(minXLineTop); // Thêm điểm nhỏ nhất phía trên vào dim LEFT
                            foreach (t3d.Point point in pointsInsight)
                            {
                                if (point.X <= minXLineTop.X + 2 && point.Y >= maxYLineLeft.Y - 2)
                                    pointsCornerLeftTop.Add(point);
                            }
                            pointsCornerLeftTop = pointsCornerLeftTop.OrderByDescending(d => d.Y).ToList(); //Sap xep theo toa do Y giam dan để chỉ lấy những điểm có tọa X bằng nhau có Y lớn hơn
                            foreach (t3d.Point point in pointsCornerLeftTop)
                            {
                                if (Math.Abs(point.X - minXLineTop.X) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerLeftTop_X.FindAll(a => Math.Abs(a.X - point.X) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerLeftTop_X.Add(point);
                                    pointsDimTop.Add(point);
                                }
                            }
                            pointsCornerLeftTop = pointsCornerLeftTop.OrderBy(d => d.X).ToList(); //Sap xep theo toa do X tăng dan để chỉ lấy những điểm có tọa X bằng nhau có X nhỏ hơn
                            foreach (t3d.Point point in pointsCornerLeftTop)
                            {
                                if (Math.Abs(point.Y - maxYLineLeft.Y) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerLeftTop_Y.FindAll(a => Math.Abs(a.Y - point.Y) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerLeftTop_Y.Add(point);
                                    pointsDimLeft.Add(point);
                                }
                            }
                        }
                        #endregion
                        #region Phần này để thêm các điểm nằm trong ô góc trên bên Phải (tạo bởi maxYLineRight và maxXLineTop)
                        if (Math.Abs(Distance.PointToLine(maxXLineTop, lineRight)) >= 3)
                        {
                            pointsDimRight.Add(maxXLineTop);//Thêm điểm X lớn nhất phía trên vào dim Right
                            pointsDimTop.Add(maxYLineRight);//Thêm điểm Y lớn nhất bên phải vào dim TOP
                            foreach (t3d.Point point in pointsInsight)
                            {
                                if (point.X >= maxXLineTop.X - 2 && point.Y >= maxYLineRight.Y - 2)
                                    pointsCornerRightTop.Add(point);
                            }
                            pointsCornerRightTop = pointsCornerRightTop.OrderByDescending(d => d.Y).ToList(); //Sap xep theo toa do Y giam dan để chỉ lấy những điểm có tọa X bằng nhau có Y lớn hơn
                            foreach (t3d.Point point in pointsCornerRightTop)
                            {
                                if (Math.Abs(point.X - maxXLineTop.X) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerRightTop_X.FindAll(a => Math.Abs(a.X - point.X) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerRightTop_X.Add(point);
                                    pointsDimTop.Add(point);
                                }
                            }
                            pointsCornerRightTop = pointsCornerRightTop.OrderByDescending(d => d.X).ToList(); //Sap xep theo toa do X giam dan để chỉ lấy những điểm có tọa Y bằng nhau có X lớn hơn
                            foreach (t3d.Point point in pointsCornerRightTop)
                            {
                                if (Math.Abs(point.Y - maxYLineRight.Y) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerRightTop_Y.FindAll(a => Math.Abs(a.Y - point.Y) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerRightTop_Y.Add(point);
                                    pointsDimRight.Add(point);
                                }
                            }
                        }
                        #endregion
                        #region Phần này để thêm các điểm nằm trong ô góc dưới bên trái (tạo bởi minYLineLeft và minXLineBot)
                        if (Math.Abs(Distance.PointToLine(minXLineBot, lineLeft)) >= 3)
                        {
                            pointsDimLeft.Add(minXLineBot);
                            pointsDimBot.Add(minYLineLeft);
                            foreach (t3d.Point point in pointsInsight)
                            {
                                if (point.X <= minXLineBot.X + 2 && point.Y <= minYLineLeft.Y + 2)
                                    pointsCornerLeftBot.Add(point);
                            }
                            pointsCornerLeftBot = pointsCornerLeftBot.OrderBy(d => d.Y).ToList(); //Sap xep theo toa do Y Tăng dan để chỉ lấy những điểm có tọa X bằng nhau có Y nhỏ hơn
                            foreach (t3d.Point point in pointsCornerLeftBot)
                            {
                                if (Math.Abs(point.X - minXLineBot.X) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerLeftBot_X.FindAll(a => Math.Abs(a.X - point.X) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerLeftBot_X.Add(point);
                                    pointsDimBot.Add(point);
                                }
                            }
                            pointsCornerLeftBot = pointsCornerLeftBot.OrderBy(p => p.X).ToList(); //Sắp xếp theo tọa độ X tăng dần để chỉ lấy những điểm có tọa Y bằng nhau có X nhỏ hơn.
                            foreach (t3d.Point point in pointsCornerLeftBot)
                            {
                                if (Math.Abs(point.Y - minYLineLeft.Y) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerLeftBot_Y.FindAll(a => Math.Abs(a.Y - point.Y) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerLeftBot_Y.Add(point);
                                    pointsDimLeft.Add(point);
                                }
                            }
                        }
                        #endregion
                        #region Phần này để thêm các điểm nằm trong ô góc dưới bên Phải (tạo bởi minYLineRight và maxXLineBot)
                        if (Math.Abs(Distance.PointToLine(maxXLineBot, lineRight)) >= 3)
                        {
                            pointsDimBot.Add(minYLineRight);
                            pointsDimRight.Add(maxXLineBot);
                            foreach (t3d.Point point in pointsInsight)
                            {
                                if (point.X >= maxXLineBot.X - 2 && point.Y <= minYLineRight.Y + 2)
                                    pointsCornerRightBot.Add(point);
                            }
                            pointsCornerRightBot = pointsCornerRightBot.OrderBy(d => d.Y).ToList(); //Sap xep theo toa do Y Tăng dan để chỉ lấy những điểm có tọa X bằng nhau có Y nhỏ hơn
                            foreach (t3d.Point point in pointsCornerRightBot)
                            {
                                if (Math.Abs(point.X - maxXLineBot.X) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerRightBot_X.FindAll(a => Math.Abs(a.X - point.X) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerRightBot_X.Add(point);
                                    pointsDimBot.Add(point);
                                }
                            }
                            pointsCornerRightBot = pointsCornerRightBot.OrderByDescending(p => p.X).ToList(); // Sắp xếp theo X giảm dần để lấy điểm có Y trùng có X lớn hơn (dim qua phải)
                            foreach (t3d.Point point in pointsCornerRightBot)
                            {
                                if (Math.Abs(point.Y - minYLineRight.Y) <= 3)
                                    continue;
                                List<t3d.Point> existing = new List<t3d.Point>();
                                existing = pointsCornerRightBot_Y.FindAll(a => Math.Abs(a.Y - point.Y) <= 3);
                                if (existing.Count == 0)
                                {
                                    pointsCornerRightBot_Y.Add(point);
                                    pointsDimRight.Add(point);
                                }
                            }
                        }
                        #endregion
                        List<t3d.Point> pointsInMid = new List<t3d.Point>(); // danh sách chứa các điểm nằm trong, không gồm các điểm thuộc 4 ô ở trên
                        pointsInMid = pointsInsight.Except(pointsCornerLeftTop).ToList();
                        pointsInMid = pointsInMid.Except(pointsCornerLeftBot).ToList();
                        pointsInMid = pointsInMid.Except(pointsCornerRightTop).ToList();
                        pointsInMid = pointsInMid.Except(pointsCornerRightBot).ToList();

                        createDimension.CreateStraightDimensionSet_X(viewbase, new PointList() { maxYLineLeft, maxYLineRight }, cb_DimentionAttributes.Text, dimSpace + dimSpaceDim2Dim + Convert.ToInt32(Math.Abs(maxYLineLeft.Y - pointsDimTop[0].Y)));
                        createDimension.CreateStraightDimensionSet_Y_(viewbase, new PointList() { minXLineTop, minXLineBot }, cb_DimentionAttributes.Text, dimSpace + dimSpaceDim2Dim + Convert.ToInt32(Math.Abs(minXLineTop.X - pointsDimLeft[0].X)));

                        if (pointsDimTop.Count > 2)
                            createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimTop, cb_DimentionAttributes.Text, dimSpace);
                        if (pointsDimBot.Count > 2)
                            createDimension.CreateStraightDimensionSet_X_(viewbase, pointsDimBot, cb_DimentionAttributes.Text, dimSpace);
                        if (pointsDimRight.Count > 2)
                            createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimRight, cb_DimentionAttributes.Text, dimSpace);
                        if (pointsDimLeft.Count > 2)
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, pointsDimLeft, cb_DimentionAttributes.Text, dimSpace);
                    }
                }
            }
            catch
            {
            }
        } //Hàm này để dim grating hoặc checker plate bị cắt khoét, được gọi trong button DIM PART OVERALL+BOLT

        public void DimTrimFlashing()
        {
            try
            {
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                DrawingHandler drawingHandler = new DrawingHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                ClearDrawingObjects clearDrawingObjects = new ClearDrawingObjects();
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                string dimAttributes = cb_DimentionAttributes.Text;
                int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drpart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                        //Lấy view hiện hành thông qua drpart
                        tsd.View viewcur = drpart.GetView() as tsd.View;
                        tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.ViewBase viewbase = drpart.GetView();
                        List<t3d.Point> listPoints = new List<t3d.Point>();
                        if (chbox_ClearCurrentDim.Checked)
                            clearDrawingObjects.ClearDim(viewcur);
                        double viewScale = viewcur.Attributes.Scale;
                        if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                        {
                            dimSpace = Convert.ToInt32(viewScale * 10);
                            dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                        }

                        //Tính điểm định nghĩa của part theo bản vẽ
                        ArrayList arlistRefLinePoint = mpart.GetReferenceLine(true);
                        t3d.Point pointStart = arlistRefLinePoint[0] as t3d.Point;
                        t3d.Point pointEnd = arlistRefLinePoint[arlistRefLinePoint.Count - 1] as t3d.Point;
                        for (int i = 0; i < arlistRefLinePoint.Count - 1; i++)
                        {
                            t3d.Point p1 = arlistRefLinePoint[i] as t3d.Point;
                            t3d.Point p2 = arlistRefLinePoint[i + 1] as t3d.Point;

                            PointList pointList = new PointList();
                            pointList.Add(p1);
                            pointList.Add(p2);
                            //MessageBox.Show(i.ToString() + "/ " + p1.ToString() + p2.ToString()); // 
                            t3d.Vector vx = new Vector(p2.X - p1.X, p2.Y - p1.Y, 0);
                            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                            if (pointStart.X > pointEnd.X)
                            {
                                StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
                                vx = new Vector(p1.X - p2.X, p1.Y - p2.Y, 0);
                                vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                try
                                {
                                    tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, -1 * vy, dimSpace, attributes);
                                    straightDimensionSet.Select();
                                    straightDimensionSet.Distance = dimSpace;
                                    straightDimensionSet.Modify();

                                }
                                catch
                                {
                                    continue;
                                }
                                //createDimension.CreateStraightDimensionSet_FX_(p1, p2, viewbase, pointList, dimAttributes, dimSpace);
                            }
                            else
                            {
                                StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
                                vx = new Vector(p1.X - p2.X, p1.Y - p2.Y, 0);
                                vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                try
                                {
                                    tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, vy, dimSpace, attributes);
                                    straightDimensionSet.Select();
                                    straightDimensionSet.Distance = dimSpace;
                                    straightDimensionSet.Modify();
                                }
                                catch
                                {
                                    continue;
                                }
                                //createDimension.CreateStraightDimensionSet_FX(p1, p2, viewbase, pointList, dimAttributes, dimSpace);
                            }
                            if (i >= 1 && i < arlistRefLinePoint.Count - 1)
                            {
                                t3d.Point p3 = arlistRefLinePoint[i - 1] as t3d.Point;
                                double angle = ((Math.Atan2(p3.Y - p1.Y, p3.X - p1.X) - Math.Atan2(p2.Y - p1.Y, p2.X - p1.X)) * 180) / 3.141592;
                                // MessageBox.Show(angle.ToString());
                                if (Math.Abs(angle - 90) < 1 || Math.Abs(angle - 270) < 0.01 || Math.Abs(angle + 90) < 1 || Math.Abs(angle + 270) < 0.01) // nếu chênh lệch với góc vuông lớn hơn 1 mới tạo
                                { }
                                else
                                {
                                    AngleDimension angleDim = new AngleDimension(viewbase, p1, p2, p3, 0);
                                    angleDim.Insert();
                                }
                            }

                        }
                    }
                    else if (drobj is tsd.View)
                    {
                        tsd.View view = drobj as tsd.View;
                        tsm.TransformationPlane viewplane = new TransformationPlane(view.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.DrawingObjectEnumerator drawingPartEnumerator = view.GetAllObjects(new Type[] { typeof(tsd.Part) });
                        foreach (tsd.DrawingObject part in drawingPartEnumerator)
                        {
                            //Convert về Part
                            tsd.Part drpart = part as tsd.Part;
                            //Lấy model part thông qua part model identifỉe
                            tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                            tsd.ViewBase viewbase = drpart.GetView();
                            List<t3d.Point> listPoints = new List<t3d.Point>();
                            ArrayList arlistRefLinePoint = mpart.GetReferenceLine(true);
                            t3d.Point pointStart = arlistRefLinePoint[0] as t3d.Point;
                            t3d.Point pointEnd = arlistRefLinePoint[arlistRefLinePoint.Count - 1] as t3d.Point;
                            for (int i = 0; i < arlistRefLinePoint.Count - 1; i++)
                            {
                                t3d.Point p1 = arlistRefLinePoint[i] as t3d.Point;
                                t3d.Point p2 = arlistRefLinePoint[i + 1] as t3d.Point;

                                PointList pointList = new PointList();
                                pointList.Add(p1);
                                pointList.Add(p2);
                                //MessageBox.Show(i.ToString() + "/ " + p1.ToString() + p2.ToString()); // 
                                t3d.Vector vx = new Vector(p2.X - p1.X, p2.Y - p1.Y, 0);
                                t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                                if (pointStart.X > pointEnd.X)
                                    createDimension.CreateStraightDimensionSet_FX_(p1, p2, viewbase, pointList, dimAttributes, dimSpace);
                                else
                                    createDimension.CreateStraightDimensionSet_FX(p1, p2, viewbase, pointList, dimAttributes, dimSpace);
                                if (i >= 1 && i < arlistRefLinePoint.Count - 1)
                                {
                                    t3d.Point p3 = arlistRefLinePoint[i - 1] as t3d.Point;
                                    double angle = ((Math.Atan2(p3.Y - p1.Y, p3.X - p1.X) - Math.Atan2(p2.Y - p1.Y, p2.X - p1.X)) * 180) / 3.141592;
                                    // MessageBox.Show(angle.ToString());
                                    if (Math.Abs(angle - 90) < 1 || Math.Abs(angle - 270) < 1 || Math.Abs(angle + 90) < 1 || Math.Abs(angle + 270) < 1) // nếu chênh lệch với góc vuông lớn hơn 1 mới tạo
                                    { }
                                    else
                                    {
                                        AngleDimension angleDim = new AngleDimension(viewbase, p1, p2, p3, 0);
                                        angleDim.Insert();
                                    }
                                }
                            }
                        }
                    }
                }
                actived_dr.CommitChanges();
            }
            catch
            { }
        }
        public void DimPickedPoints()
        {
            try
            {
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                DrawingHandler drawingHandler = new DrawingHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                ClearDrawingObjects clearDrawingObjects = new ClearDrawingObjects();
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                string dimAttributes = cb_DimentionAttributes.Text;
                int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
                tsdui.Picker picker = drawingHandler.GetPicker();
                StringList stringList = new StringList();
                stringList.Add("Pick points");
                Tuple<PointList, ViewBase> tuple = picker.PickPoints(stringList);
                PointList pointList1 = tuple.Item1;
                tsd.ViewBase viewbase = tuple.Item2;
                tsd.View viewcur = viewbase as tsd.View;
                tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                wph.SetCurrentTransformationPlane(viewplane);
                if (chbox_ClearCurrentDim.Checked)
                    clearDrawingObjects.ClearDim(viewcur);
                double viewScale = viewcur.Attributes.Scale;
                if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                {
                    dimSpace = Convert.ToInt32(viewScale * 10);
                    dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                }
                ArrayList arlistRefLinePoint = new ArrayList();
                foreach (var item in pointList1)
                {
                    arlistRefLinePoint.Add(item);
                }
                t3d.Point pointStart = arlistRefLinePoint[0] as t3d.Point;
                t3d.Point pointEnd = arlistRefLinePoint[arlistRefLinePoint.Count - 1] as t3d.Point;
                for (int i = 0; i < arlistRefLinePoint.Count - 2; i++) // cho nay sua cho tekla 2020 la -1
                {
                    t3d.Point p1 = arlistRefLinePoint[i] as t3d.Point;
                    t3d.Point p2 = arlistRefLinePoint[i + 1] as t3d.Point;

                    PointList pointList = new PointList();
                    pointList.Add(p1);
                    pointList.Add(p2);
                    //MessageBox.Show(i.ToString() + "/ " + p1.ToString() + p2.ToString()); // 
                    t3d.Vector vx = new Vector(p2.X - p1.X, p2.Y - p1.Y, 0);
                    t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                    try
                    {
                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
                            vx = new Vector(p1.X - p2.X, p1.Y - p2.Y, 0);
                            vy = vx.Cross(new t3d.Vector(0, 0, 1));
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, vy, dimSpace, attributes);
                                straightDimensionSet.Select();
                                straightDimensionSet.Distance = dimSpace;
                                straightDimensionSet.Modify();
                            }
                            else
                            {
                                tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, -1 * vy, dimSpace, attributes);
                                straightDimensionSet.Select();
                                straightDimensionSet.Distance = dimSpace;
                                straightDimensionSet.Modify();
                            }
                        }
                        else
                        {
                            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
                            vx = new Vector(p2.X - p1.X, p2.Y - p1.Y, 0);
                            vy = vx.Cross(new t3d.Vector(0, 0, 1));
                            if (Control.ModifierKeys == Keys.Shift)
                            {
                                tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, -1 * vy, dimSpace, attributes);
                                straightDimensionSet.Select();
                                straightDimensionSet.Distance = dimSpace;
                                straightDimensionSet.Modify();
                            }
                            else
                            {
                                tsd.StraightDimensionSet straightDimensionSet = straightDimensionSetHandler.CreateDimensionSet(viewbase, pointList, vy, dimSpace, attributes);
                                straightDimensionSet.Select();
                                straightDimensionSet.Distance = dimSpace;
                                straightDimensionSet.Modify();
                            }
                            //createDimension.CreateStraightDimensionSet_FX(p1, p2, viewbase, pointList, dimAttributes, dimSpace);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                    if (i >= 1 && i < arlistRefLinePoint.Count - 1)
                    {
                        t3d.Point p3 = arlistRefLinePoint[i - 1] as t3d.Point;
                        double angle = ((Math.Atan2(p3.Y - p1.Y, p3.X - p1.X) - Math.Atan2(p2.Y - p1.Y, p2.X - p1.X)) * 180) / 3.141592;
                        // MessageBox.Show(angle.ToString());
                        if (Math.Abs(angle - 90) < 1 || Math.Abs(angle - 270) < 0.01 || Math.Abs(angle + 90) < 1 || Math.Abs(angle + 270) < 0.01) // nếu chênh lệch với góc vuông lớn hơn 1 mới tạo
                        { }
                        else
                        {
                            AngleDimension angleDim = new AngleDimension(viewbase, p1, p2, p3, 0);
                            angleDim.Insert();
                        }
                    }
                }
                actived_dr.CommitChanges();
            }
            catch
            { }
        }
        public void DimCurveTole()
        {
            try
            {
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                DrawingHandler drawingHandler = new DrawingHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                ClearDrawingObjects clearDrawingObjects = new ClearDrawingObjects();
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                string dimAttributes = cb_DimentionAttributes.Text;
                int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drpart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                        //Lấy view hiện hành thông qua drpart
                        tsd.View view = drpart.GetView() as tsd.View;
                        tsm.TransformationPlane viewplane = new TransformationPlane(view.DisplayCoordinateSystem);
                        if (chbox_ClearCurrentDim.Checked)
                            clearDrawingObjects.ClearDim(view);
                        double viewScale = view.Attributes.Scale;
                        if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                        {
                            dimSpace = Convert.ToInt32(viewScale * 10);
                            dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                        }
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.ViewBase viewbase = drpart.GetView();
                        List<t3d.Point> listPoints = new List<t3d.Point>();
                        ArrayList arlistRefLinePoint = mpart.GetReferenceLine(true);
                        //MessageBox.Show(arlistRefLinePoint.Count.ToString());
                        t3d.Point p1 = arlistRefLinePoint[0] as t3d.Point;
                        t3d.Point p2 = arlistRefLinePoint[1] as t3d.Point;
                        t3d.Point p3 = arlistRefLinePoint[(arlistRefLinePoint.Count - 1) / 2] as t3d.Point;
                        t3d.Point p4 = arlistRefLinePoint[arlistRefLinePoint.Count - 2] as t3d.Point;
                        t3d.Point p5 = arlistRefLinePoint[arlistRefLinePoint.Count - 1] as t3d.Point;

                        t3d.Line line1 = new t3d.Line(p1, p2);
                        t3d.Line line2 = new t3d.Line(p4, p5);
                        t3d.LineSegment lineSegment = Intersection.LineToLine(line1, line2);

                        t3d.Point pointMid = lineSegment.Point1;
                        double distanceAngleDim = Distance.PointToPoint(p4, pointMid);


                        PointList plDimXYF_total = new PointList(); //ds điểm p1 p5
                        PointList plDimY_total = new PointList(); //ds điểm p1 p5
                        //PointList plDimF_total = new PointList(); //ds điểm p1 p5
                        PointList plDimF1 = new PointList(); //ds điểm p1 p2
                        PointList plDimF2 = new PointList(); //ds điểm p4 p5
                        PointList plCurveLength = new PointList(); //ds điểm dim cung đơn vị mm
                        PointList plCurveRadial = new PointList(); //ds điểm dim cung đơn vị độ
                        PointList plRadial = new PointList();//ds điểm dim bán kính R


                        plDimXYF_total.Add(p5);
                        plDimXYF_total.Add(p1);
                        plDimY_total.Add(p5);
                        plDimY_total.Add(p1);
                        plDimF1.Add(p1);
                        plDimF1.Add(p2);
                        plDimF2.Add(p4);
                        plDimF2.Add(p5);

                        plRadial.Add(p2);
                        //plRadial.Add(p3);
                        plRadial.Add(p4);

                        //MessageBox.Show(p1.ToString() + "-" + p2.ToString() + "-" + p3.ToString() + "-" + p4.ToString() + "-" + p5.ToString());
                        AngleDimension myAngle = new AngleDimension(viewbase, pointMid, p4, p2, dimSpaceDim2Dim);
                        myAngle.Attributes = new AngleDimensionAttributes("AngleDim");
                        myAngle.Attributes.Type = AngleTypes.AngleAtVertex;
                        myAngle.Insert();

                        CurvedDimensionSetRadial aaa = new CurvedDimensionSetHandler().CreateCurvedDimensionSetRadial(viewbase, p2, p3, p4, plRadial, dimSpace + 1000); // phải tạo 1 dim curve này thì các dim ở code phía dưới mới đúng điểm được.
                        aaa.Delete();

                        new CurvedDimensionSetHandler().CreateCurvedDimensionSetRadial(viewbase, p2, p3, p4, plRadial, -0.1 * dimSpace);
                        RadiusDimension radiusDimension = new RadiusDimension(viewbase, p2, p3, p4, -10);
                        radiusDimension.Insert();
                        radiusDimension.Select();
                        radiusDimension.Distance = -10;
                        radiusDimension.Modify();

                        StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
                        tsd.StraightDimensionSet straightDimensionSetXTotal = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimXYF_total, new t3d.Vector(0, -1, 0), 0, attributes);
                        straightDimensionSetXTotal.Select();
                        straightDimensionSetXTotal.Distance = dimSpace + Convert.ToInt32(Math.Abs(p1.Y - p5.Y));
                        straightDimensionSetXTotal.Modify();

                        tsd.StraightDimensionSet straightDimensionSetYTotal = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimY_total, new t3d.Vector(-1, 0, 0), 0, attributes);
                        straightDimensionSetYTotal.Select();
                        straightDimensionSetYTotal.Distance = dimSpace + Convert.ToInt32(Math.Abs(p1.X - p5.X));
                        straightDimensionSetYTotal.Modify();

                        t3d.Vector vx = new Vector(p1.X - p5.X, p1.Y - p5.Y, 0);
                        t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                        tsd.StraightDimensionSet straightDimensionSetFX_ = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimXYF_total, -1 * vy, dimSpace, attributes);
                        straightDimensionSetFX_.Select();
                        straightDimensionSetFX_.Distance = dimSpace;
                        straightDimensionSetFX_.Modify();

                        tsd.StraightDimensionSet straightDimensionSetX1_ = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimF1, new t3d.Vector(0, -1, 0), 0, attributes);
                        straightDimensionSetX1_.Select();
                        straightDimensionSetX1_.Distance = dimSpace;
                        straightDimensionSetX1_.Modify();

                        t3d.Vector vx2 = new Vector(p4.X - p5.X, p4.Y - p5.Y, 0);
                        t3d.Vector vy2 = vx2.Cross(new t3d.Vector(0, 0, 1));
                        tsd.StraightDimensionSet straightDimensionSetFX2_ = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimF2, -1 * vy2, dimSpace, attributes);
                        straightDimensionSetFX2_.Select();
                        straightDimensionSetFX2_.Distance = dimSpace;
                        straightDimensionSetFX2_.Modify();
                        //createDimension.CreateStraightDimensionSet_X_(viewbase, plDimXYF_total, dimAttributes, dimSpace );
                        //createDimension.CreateStraightDimensionSet_Y_(viewbase, plDimY_total, dimAttributes, dimSpace);
                        //createDimension.CreateStraightDimensionSet_FX_(p1, p5, viewbase, plDimXYF_total, dimAttributes, dimSpace);
                        //createDimension.CreateStraightDimensionSet_X_(viewbase, plDimF1, dimAttributes, dimSpace);
                        //createDimension.CreateStraightDimensionSet_FX(p5, p4, viewbase, plDimF2, dimAttributes, dimSpace);
                    }
                    else if (drobj is tsd.View)
                    {
                        tsd.View view = drobj as tsd.View;
                        if (chbox_ClearCurrentDim.Checked)
                            clearDrawingObjects.ClearDim(view);

                        tsm.TransformationPlane viewplane = new TransformationPlane(view.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.DrawingObjectEnumerator drawingPartEnumerator = view.GetAllObjects(new Type[] { typeof(tsd.Part) });
                        foreach (tsd.DrawingObject part in drawingPartEnumerator)
                        {
                            //Convert về Part
                            tsd.Part drpart = part as tsd.Part;
                            //Lấy model part thông qua part model identifỉe
                            tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                            tsd.ViewBase viewbase = drpart.GetView();
                            List<t3d.Point> listPoints = new List<t3d.Point>();
                            ArrayList arlistRefLinePoint = mpart.GetReferenceLine(true);
                            //MessageBox.Show(arlistRefLinePoint.Count.ToString());
                            t3d.Point p1 = arlistRefLinePoint[0] as t3d.Point;
                            t3d.Point p2 = arlistRefLinePoint[1] as t3d.Point;
                            t3d.Point p3 = arlistRefLinePoint[(arlistRefLinePoint.Count - 1) / 2] as t3d.Point;
                            t3d.Point p4 = arlistRefLinePoint[arlistRefLinePoint.Count - 2] as t3d.Point;
                            t3d.Point p5 = arlistRefLinePoint[arlistRefLinePoint.Count - 1] as t3d.Point;

                            t3d.Line line1 = new t3d.Line(p1, p2);
                            t3d.Line line2 = new t3d.Line(p4, p5);
                            t3d.LineSegment lineSegment = Intersection.LineToLine(line1, line2);

                            t3d.Point pointMid = lineSegment.Point1;
                            double distanceAngleDim = Distance.PointToPoint(p4, pointMid);


                            PointList plDimXYF_total = new PointList(); //ds điểm p1 p5
                            PointList plDimY_total = new PointList(); //ds điểm p1 p5
                                                                      //PointList plDimF_total = new PointList(); //ds điểm p1 p5
                            PointList plDimF1 = new PointList(); //ds điểm p1 p2
                            PointList plDimF2 = new PointList(); //ds điểm p4 p5
                            PointList plCurveLength = new PointList(); //ds điểm dim cung đơn vị mm
                            PointList plCurveRadial = new PointList(); //ds điểm dim cung đơn vị độ
                            PointList plRadial = new PointList();//ds điểm dim bán kính R


                            plDimXYF_total.Add(p1);
                            plDimXYF_total.Add(p5);
                            plDimY_total.Add(p5);
                            plDimY_total.Add(p1);
                            plDimF1.Add(p1);
                            plDimF1.Add(p2);
                            plDimF2.Add(p4);
                            plDimF2.Add(p5);

                            plRadial.Add(p2);
                            //plRadial.Add(p3);
                            plRadial.Add(p4);

                            //MessageBox.Show(p1.ToString() + "-" + p2.ToString() + "-" + p3.ToString() + "-" + p4.ToString() + "-" + p5.ToString());
                            AngleDimension myAngle = new AngleDimension(viewbase, pointMid, p4, p2, dimSpaceDim2Dim);
                            myAngle.Attributes = new AngleDimensionAttributes("AngleDim");
                            myAngle.Attributes.Type = AngleTypes.AngleAtVertex;
                            myAngle.Insert();

                            CurvedDimensionSetRadial aaa = new CurvedDimensionSetHandler().CreateCurvedDimensionSetRadial(viewbase, p2, p3, p4, plRadial, dimSpace + 1000); // phải tạo 1 dim curve này thì các dim ở code phía dưới mới đúng điểm được.
                            aaa.Delete();

                            new CurvedDimensionSetHandler().CreateCurvedDimensionSetRadial(viewbase, p2, p3, p4, plRadial, -0.1 * dimSpace);
                            RadiusDimension radiusDimension = new RadiusDimension(viewbase, p2, p3, p4, -10);
                            radiusDimension.Insert();
                            radiusDimension.Select();
                            radiusDimension.Distance = -10;
                            radiusDimension.Modify();

                            StraightDimensionSet.StraightDimensionSetAttributes attributes = new StraightDimensionSet.StraightDimensionSetAttributes(null, dimAttributes);
                            tsd.StraightDimensionSet straightDimensionSetXTotal = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimXYF_total, new t3d.Vector(0, -1, 0), 0, attributes);
                            straightDimensionSetXTotal.Select();
                            straightDimensionSetXTotal.Distance = dimSpace + Convert.ToInt32(Math.Abs(p1.Y - p5.Y));
                            straightDimensionSetXTotal.Modify();

                            tsd.StraightDimensionSet straightDimensionSetYTotal = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimY_total, new t3d.Vector(-1, 0, 0), 0, attributes);
                            straightDimensionSetYTotal.Select();
                            straightDimensionSetYTotal.Distance = dimSpace + Convert.ToInt32(Math.Abs(p1.X - p5.X));
                            straightDimensionSetYTotal.Modify();

                            t3d.Vector vx = new Vector(p1.X - p5.X, p1.Y - p5.Y, 0);
                            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                            tsd.StraightDimensionSet straightDimensionSetFX_ = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimXYF_total, -1 * vy, dimSpace, attributes);
                            straightDimensionSetFX_.Select();
                            straightDimensionSetFX_.Distance = dimSpace;
                            straightDimensionSetFX_.Modify();

                            tsd.StraightDimensionSet straightDimensionSetX1_ = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimF1, new t3d.Vector(0, -1, 0), 0, attributes);
                            straightDimensionSetX1_.Select();
                            straightDimensionSetX1_.Distance = dimSpace;
                            straightDimensionSetX1_.Modify();

                            t3d.Vector vx2 = new Vector(p4.X - p5.X, p4.Y - p5.Y, 0);
                            t3d.Vector vy2 = vx2.Cross(new t3d.Vector(0, 0, 1));
                            tsd.StraightDimensionSet straightDimensionSetFX2_ = straightDimensionSetHandler.CreateDimensionSet(viewbase, plDimF2, -1 * vy2, dimSpace, attributes);
                            straightDimensionSetFX2_.Select();
                            straightDimensionSetFX2_.Distance = dimSpace;
                            straightDimensionSetFX2_.Modify();
                            //createDimension.CreateStraightDimensionSet_X_(viewbase, plDimXYF_total, dimAttributes, dimSpace );
                            //createDimension.CreateStraightDimensionSet_Y_(viewbase, plDimY_total, dimAttributes, dimSpace);
                            //createDimension.CreateStraightDimensionSet_FX_(p1, p5, viewbase, plDimXYF_total, dimAttributes, dimSpace);
                            //createDimension.CreateStraightDimensionSet_X_(viewbase, plDimF1, dimAttributes, dimSpace);
                            //createDimension.CreateStraightDimensionSet_FX(p5, p4, viewbase, plDimF2, dimAttributes, dimSpace);

                        }
                    }
                }
                actived_dr.CommitChanges();
            }
            catch
            { }
        }

        public void DimPurlin()
        {
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            wph.SetCurrentTransformationPlane(new TransformationPlane());
            CreateDimension createDimension = new CreateDimension();
            ClearDrawingObjects clearDrawingObjects = new ClearDrawingObjects();
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            tsdui.Picker c_picker = drawingHandler.GetPicker();
            string dimAttribute = cb_DimentionAttributes.Text; //Khoảng cb_DimentionAttributes mặc định
            int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
            int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
            Function function = new Function();
            double offsetFromFrame = 5;
            double minCutpartDecrease = 20;
            ResizeView resizeView = new ResizeView();

            tsd.PointList Boltlist_Bolt_XY = new PointList();
            List<t3d.Point> Boltlist_Bolt_XY_Only = new List<t3d.Point>(); //để tính toán ra điểm tạo line giữa các point của BoltXY (Boltlist_Bolt_XY bị add thêm các point k cần thiết)
            tsd.PointList Boltlist_Bolt_Y_top = new PointList(); //ds các điểm bolt Y nằm trên center line
            tsd.PointList Boltlist_Bolt_Y_bot = new PointList(); //ds các điểm bolt Y nằm dưới center line
            tsd.PointList Boltlist_Bolt_X = new PointList();
            tsd.PointList Boltlist_Bolt_skew = new PointList();
            tsd.PointList dim_Bolt_Y_top = new PointList(); //ds các điểm bolt Y nằm trên center line
            tsd.PointList dim_Bolt_Y_bot = new PointList(); //ds các điểm bolt Y nằm dưới center line
            tsd.PointList dim_Bolt_X = new PointList();
            tsd.PointList dimOverall_X = new PointList();
            tsd.PointList dimOverall_X_ = new PointList();
            tsd.PointList dimOverall_Y = new PointList();
            tsd.PointList dimOverall_Y1 = new PointList();
            tsd.PointList dimOverall_Y2 = new PointList();

            foreach (tsd.DrawingObject drobj in drobjenum)
            {
                if (drobj is tsd.Part)
                {
                    //Convert đối tượng về part
                    tsd.Part drPart = drobj as tsd.Part;
                    //chuyển dr_part về model
                    tsm.Part modelPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                    tsd.View view = drPart.GetView() as tsd.View;
                    //khai báo viewbase
                    tsd.ViewBase viewBase = view as tsd.ViewBase;
                    if (chbox_ClearCurrentDim.Checked)
                        clearDrawingObjects.ClearDim(view);
                    double viewScale = view.Attributes.Scale;
                    if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                    {
                        dimSpace = Convert.ToInt32(viewScale * 3 * 5);
                        dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                    }
                    string partProfileType = function.GetProfileType(modelPart);
                    //tính toán điểm của part này trên view bản vẽ
                    Part_Edge partEdge = new Part_Edge(view, modelPart); //tạo 1 myPartEdge thuộc kiểu dữ liệu (class) Part_Edge
                    //model.CommitChanges();
                    List<t3d.Point> pointListPart = partEdge.List_Edge;
                    ArrayList listPointCenter = modelPart.GetCenterLine(true);
                    t3d.Point clPoint1 = listPointCenter[0] as t3d.Point;
                    t3d.Point clPoint2 = listPointCenter[1] as t3d.Point;
                    t3d.Point mid_YofPart = partEdge.PointMidLeft;
                    t3d.Point pointXmaxYmax = partEdge.PointXmaxYmax;
                    t3d.Point pointXmaxYmin = partEdge.PointXmaxYmin;
                    t3d.Point pointXminYmax = partEdge.PointXminYmax;
                    t3d.Point pointXminYmin = partEdge.PointXminYmin;

                    t3d.Point pointXmin0 = partEdge.PointXmin0;
                    t3d.Point pointXmin1 = partEdge.PointXmin1;
                    t3d.Point pointXmax0 = partEdge.PointXmax0;
                    t3d.Point pointXmax1 = partEdge.PointXmax1;
                    t3d.Point pointYmin0 = partEdge.PointYmin0;
                    t3d.Point pointYmin1 = partEdge.PointYmin1;
                    t3d.Point pointYmax0 = partEdge.PointYmax0;
                    t3d.Point pointYmax1 = partEdge.PointYmax1;

                    dimOverall_X.Add(pointXminYmax);
                    dimOverall_X.Add(pointXmaxYmax);
                    dimOverall_Y.Add(pointXminYmax);
                    dimOverall_Y.Add(pointXminYmin);

                    Boltlist_Bolt_XY.Add(pointXminYmax);
                    Boltlist_Bolt_XY.Add(pointXmaxYmax);
                    Boltlist_Bolt_XY.Add(pointXminYmin);

                    Boltlist_Bolt_Y_top.Add(pointXminYmax);
                    Boltlist_Bolt_Y_top.Add(pointXmaxYmax);
                    Boltlist_Bolt_Y_bot.Add(pointXminYmin);
                    Boltlist_Bolt_Y_bot.Add(pointXmaxYmin);
                    tsm.ModelObjectEnumerator boltenum = modelPart.GetBolts();

                    foreach (tsm.BoltGroup bgr in boltenum)
                    {
                        Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                        //Tìm bolt theo phương XY
                        tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY;
                        tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                        //wph.SetCurrentTransformationPlane(viewplane);
                        if (bolt_xy != null)
                        {
                            bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                            foreach (t3d.Point p in bolt_xy.BoltPositions)
                            {
                                Boltlist_Bolt_XY.Add(p);
                                Boltlist_Bolt_XY_Only.Add(p);
                            }
                        }
                        if (bolt_y != null)
                        {
                            bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                            foreach (t3d.Point p in bolt_y.BoltPositions)
                            {
                                if (p.Y > mid_YofPart.Y)
                                {
                                    Boltlist_Bolt_Y_top.Add(p);
                                }
                                else
                                {
                                    Boltlist_Bolt_Y_bot.Add(p);
                                }
                            }
                        }
                    }

                    boltenum.Reset();
                    //Tính toán điểm tạo line
                    if (chbox_AddLine.Checked)
                    {
                        //tính toán các điểm để vẽ line
                        List<t3d.Point> PointListOrderBy_X = Boltlist_Bolt_XY_Only.OrderBy(point => point.X).ToList(); //để Lấy Xmin0 và Xmin1 (nếu có) 
                        t3d.Point pointBoltXY_Xmin0 = null; // là điểm X có tọa độ Y nhỏ hơn
                        t3d.Point pointBoltXY_Xmin1 = null;// là điểm X có tọa độ Y lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                        if (Math.Abs(PointListOrderBy_X[0].X - PointListOrderBy_X[1].X) <= 0.01) // nếu mà 2 điểm nhỏ nhất trong danh sách điểm sắp xếp X tăng dần có tọa độ X bằng nhau
                        {
                            if (PointListOrderBy_X[0].Y > PointListOrderBy_X[1].Y)// Nếu tọa độ Y của điểm 0 lớn hơn điểm 1 thì điểm 0 là Xmin1, điểm 1 là Xmin0 và ngược lại
                            {
                                pointBoltXY_Xmin0 = PointListOrderBy_X[1];
                                pointBoltXY_Xmin1 = PointListOrderBy_X[0];
                            }
                            else if (PointListOrderBy_X[0].Y < PointListOrderBy_X[1].Y)
                            {
                                pointBoltXY_Xmin0 = PointListOrderBy_X[0];
                                pointBoltXY_Xmin1 = PointListOrderBy_X[1];
                            }
                        }
                        else pointBoltXY_Xmin0 = PointListOrderBy_X[0];
                        //MessageBox.Show(pointBoltXY_Xmin0.ToString() + "_" + pointBoltXY_Xmin1.ToString());
                        List<t3d.Point> PointListOrderByDescending_X = Boltlist_Bolt_XY_Only.OrderByDescending(point => point.X).ToList();//để Lấy Xmax0 và Xmax1 (nếu có) 
                        t3d.Point pointBoltXY_Xmax0 = null; // là điểm Xmax có tọa độ Y nhỏ hơn
                        t3d.Point pointBoltXY_Xmax1 = null; // là điểm Xmax có tọa độ Y lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                        if (Math.Abs(PointListOrderByDescending_X[0].X - PointListOrderByDescending_X[1].X) <= 0.01) // nếu mà 2 điểm lớn nhất trong danh sách điểm sắp xếp X giảm dần có tọa độ X bằng nhau
                        {
                            if (PointListOrderByDescending_X[0].Y > PointListOrderByDescending_X[1].Y)
                            {
                                pointBoltXY_Xmax0 = PointListOrderByDescending_X[1];
                                pointBoltXY_Xmax1 = PointListOrderByDescending_X[0];
                            }
                            else if (PointListOrderByDescending_X[0].Y < PointListOrderByDescending_X[1].Y)
                            {
                                pointBoltXY_Xmax0 = PointListOrderByDescending_X[0];
                                pointBoltXY_Xmax1 = PointListOrderByDescending_X[1];
                            }
                        }
                        else pointXmax0 = PointListOrderByDescending_X[0];

                        List<t3d.Point> PointListOrderBy_Y = Boltlist_Bolt_XY_Only.OrderBy(point => point.Y).ToList();              //Lấy minY
                        t3d.Point pointBoltXY_Ymin0 = null; // là điểm Ymin có tọa độ X nhỏ hơn
                        t3d.Point pointBoltXY_Ymin1 = null;// là điểm Ymin có tọa độ X lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                        if (Math.Abs(PointListOrderBy_Y[0].Y - PointListOrderBy_Y[1].Y) <= 0.01) // nếu mà 2 điểm nhỏ nhất trong danh sách điểm sắp xếp Y tăng dần có tọa độ Y bằng nhau
                        {
                            if (PointListOrderBy_Y[0].X > PointListOrderBy_Y[1].X)// Nếu tọa độ X của điểm 0 lớn hơn điểm 1 thì điểm 0 là Ymin1, điểm 1 là Ymin0 và ngược lại
                            {
                                pointBoltXY_Ymin0 = PointListOrderBy_Y[1];
                                pointBoltXY_Ymin1 = PointListOrderBy_Y[0];
                            }
                            else if (PointListOrderBy_Y[0].X < PointListOrderBy_Y[1].X)
                            {
                                pointBoltXY_Ymin0 = PointListOrderBy_Y[0];
                                pointBoltXY_Ymin1 = PointListOrderBy_Y[1];
                            }
                        }
                        else pointYmin0 = PointListOrderBy_Y[0];

                        List<t3d.Point> PointListOrderByDescending_Y = Boltlist_Bolt_XY_Only.OrderByDescending(point => point.Y).ToList();    //Lấy maxY
                        t3d.Point pointBoltXY_Ymax0 = null; // là điểm Ymin có tọa độ X nhỏ hơn
                        t3d.Point pointBoltXY_Ymax1 = null;// là điểm Ymin có tọa độ X lớn hơn nếu có 2 điểm cùng tọa độ X (Tối đa chỉ có 2 điểm thôi)
                        if (Math.Abs(PointListOrderByDescending_Y[0].Y - PointListOrderByDescending_Y[1].Y) <= 0.01) // nếu mà 2 điểm nhỏ nhất trong danh sách điểm sắp xếp Y tăng dần có tọa độ Y bằng nhau
                        {
                            if (PointListOrderByDescending_Y[0].X > PointListOrderByDescending_Y[1].X)// Nếu tọa độ X của điểm 0 lớn hơn điểm 1 thì điểm 0 là Ymin1, điểm 1 là Ymin0 và ngược lại
                            {
                                pointBoltXY_Ymax0 = PointListOrderByDescending_Y[1];
                                pointBoltXY_Ymax1 = PointListOrderByDescending_Y[0];
                            }
                            else if (PointListOrderByDescending_Y[0].X < PointListOrderByDescending_Y[1].X)
                            {
                                pointBoltXY_Ymax0 = PointListOrderByDescending_Y[0];
                                pointBoltXY_Ymax1 = PointListOrderByDescending_Y[1];
                            }
                        }
                        else pointYmax0 = PointListOrderByDescending_Y[0];
                        tsd.Line line1 = new tsd.Line(view, pointBoltXY_Xmin0, pointBoltXY_Xmax0);
                        line1.Attributes.Line.Color = DrawingColors.Red;
                        line1.Attributes.Line.Type = LineTypes.DashDot;
                        tsd.Line line2 = new tsd.Line(view, pointBoltXY_Xmin1, pointBoltXY_Xmax1);
                        line2.Attributes.Line.Color = DrawingColors.Red;
                        line2.Attributes.Line.Type = LineTypes.DashDot;
                        line1.Insert();
                        line2.Insert();
                    }
                    //Tất cả những liên quan đến khởi tạo dim sẽ có câu Handler
                    tsd.StraightDimensionSetHandler smyDrawingHandler = new StraightDimensionSetHandler();
                    //Nếu 2 điểm center line trùng nhau thì là dim section
                    if (Math.Abs(clPoint1.X - clPoint2.X) < 3)
                    {
                        //Trường hợp Xà gồ Z móc trên quay qua bên Trái
                        if (pointXmin1 == pointYmax0 && partProfileType == "Z")
                        {
                            // MessageBox.Show("Quay trai");
                            dimOverall_X.Clear();
                            dimOverall_X_.Clear();
                            dimOverall_Y.Clear();
                            dimOverall_X.Add(pointXmin1);
                            dimOverall_X.Add(pointYmax1);
                            dimOverall_X_.Add(pointYmin0);
                            dimOverall_X_.Add(pointYmin1);
                            dimOverall_Y.Add(pointXmin1);
                            dimOverall_Y.Add(pointYmin0);
                            dimOverall_Y1.Add(pointXmin1);
                            dimOverall_Y1.Add(pointXmin0);
                            dimOverall_Y2.Add(pointXmax1);
                            dimOverall_Y2.Add(pointXmax0);
                            //tạo dim X lên trên
                            createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim X xuống dưới
                            createDimension.CreateStraightDimensionSet_X_(viewBase, dimOverall_X_, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim Y qua trái
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y1, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y2, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            dim_Bolt_Y_top.Add(pointXmin1);
                            dim_Bolt_Y_top.Add(pointYmax1);
                            dim_Bolt_Y_bot.Add(pointYmin0);
                            dim_Bolt_Y_bot.Add(pointYmin1);
                            dim_Bolt_X.Add(pointXmin1);
                            dim_Bolt_X.Add(pointYmin0);
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                //wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_x != null)
                                {
                                    bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_x.BoltPositions)
                                        dim_Bolt_X.Add(p);
                                }
                                if (bolt_y != null)
                                {
                                    bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_y.BoltPositions)
                                    {
                                        if (p.Y > mid_YofPart.Y)
                                            dim_Bolt_Y_top.Add(p);
                                        else
                                            dim_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                            if (dim_Bolt_Y_top.Count != 2)
                                createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_Y_bot.Count != 2)
                                createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_X.Count != 2)
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        }
                        //Trường hợp Xà gồ Z móc trên quay qua bên Phải
                        else if (pointXmax1 == pointYmax1 && partProfileType == "Z")
                        {
                            dimOverall_X.Clear();
                            dimOverall_X_.Clear();
                            dimOverall_Y.Clear();
                            dimOverall_X.Add(pointYmax0);
                            dimOverall_X.Add(pointYmax1);
                            dimOverall_X_.Add(pointXmin0);
                            dimOverall_X_.Add(pointYmin1);
                            dimOverall_Y.Add(pointYmax1);
                            dimOverall_Y.Add(pointYmin1);
                            dimOverall_Y1.Add(pointXmax1);
                            dimOverall_Y1.Add(pointXmax0);
                            dimOverall_Y2.Add(pointXmin1);
                            dimOverall_Y2.Add(pointXmin0);
                            //tạo dim X lên trên
                            createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim X xuống dưới
                            createDimension.CreateStraightDimensionSet_X_(viewBase, dimOverall_X_, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim Y qua trái
                            createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y1, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y2, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            dim_Bolt_Y_top.Add(pointYmax0);
                            dim_Bolt_Y_top.Add(pointYmax1);
                            dim_Bolt_Y_bot.Add(pointXmin0);
                            dim_Bolt_Y_bot.Add(pointYmin1);
                            dim_Bolt_X.Add(pointYmax1);
                            dim_Bolt_X.Add(pointYmin1);
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                //wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_x != null)
                                {
                                    bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_x.BoltPositions)
                                        dim_Bolt_X.Add(p);
                                }
                                if (bolt_y != null)
                                {
                                    bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_y.BoltPositions)
                                    {
                                        if (p.Y > mid_YofPart.Y)
                                            dim_Bolt_Y_top.Add(p);
                                        else
                                            dim_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                            if (dim_Bolt_Y_top.Count != 2)
                                createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_Y_bot.Count != 2)
                                createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_X.Count != 2)
                                createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        }
                        //Trường hợp xà gồ C
                        else if (partProfileType == "C")
                        {
                            //tạo ds các điểm có cùng tọa độ X min, Nếu số phần từ là 4 thì C quay lưng về bên trái và ngược lại.
                            List<t3d.Point> pointListPart_minX = pointListPart.FindAll(p => Math.Abs(p.X - pointXmin0.X) < 0.1).ToList();
                            //C quay lung qua phai

                            dimOverall_X.Add(pointYmax0);
                            dimOverall_X.Add(pointYmax1);
                            //Nếu C quay lưng qua trái
                            if (pointListPart_minX.Count == 4)
                            {
                                dimOverall_Y.Add(pointYmax1);
                                dimOverall_Y.Add(pointYmin1);
                            }
                            //Nếu C quay lưng qua phải
                            else
                            {
                                dimOverall_Y.Add(pointYmax0);
                                dimOverall_Y.Add(pointYmin0);
                            }
                            dim_Bolt_Y_top.Add(pointYmax0);
                            dim_Bolt_Y_top.Add(pointYmax1);
                            dim_Bolt_Y_bot.Add(pointYmin0);
                            dim_Bolt_Y_bot.Add(pointYmin1);
                            //Nếu C quay lưng qua trái
                            if (pointListPart_minX.Count == 4)
                            {
                                dim_Bolt_X.Add(pointYmax1);
                                dim_Bolt_X.Add(pointYmin1);
                            }
                            //Nếu C quay lưng qua phải
                            else
                            {
                                dim_Bolt_X.Add(pointYmax0);
                                dim_Bolt_X.Add(pointYmin0);
                            }
                            //tạo dim X lên trên
                            createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim Y qua trái
                            if (pointListPart_minX.Count == 4)
                                createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                            else //tạo dim Y qua phải
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                //wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_x != null)
                                {
                                    bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_x.BoltPositions)
                                        dim_Bolt_X.Add(p);
                                }
                                if (bolt_y != null)
                                {
                                    bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_y.BoltPositions)
                                    {
                                        if (p.Y > mid_YofPart.Y)
                                            dim_Bolt_Y_top.Add(p);
                                        else
                                            dim_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                            if (dim_Bolt_Y_top.Count != 2)
                                createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_Y_bot.Count != 2)
                                createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim);

                            if (dim_Bolt_X.Count != 2)
                            {
                                if (pointListPart_minX.Count == 4)
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                else //tạo dim Y qua phải
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            }
                        }
                    }
                    //Nếu không sẽ dim purlin bình thường
                    else
                    {
                        //Tạo dimension bao theo phương X
                        createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace);
                        //Tạo dimension bao theo phương Y
                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                        //Tạo bolt xy theo phương X
                        createDimension.CreateStraightDimensionSet_X(viewBase, Boltlist_Bolt_XY, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        //Tạo bolt xy theo phương Y
                        createDimension.CreateStraightDimensionSet_Y_(viewBase, Boltlist_Bolt_XY, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        //Tạo bolt Y theo phương X
                        if (Boltlist_Bolt_Y_top.Count > 2)
                        {
                            createDimension.CreateStraightDimensionSet_X(viewBase, Boltlist_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                        }
                        if (Boltlist_Bolt_Y_bot.Count > 2)
                        {
                            createDimension.CreateStraightDimensionSet_X_(viewBase, Boltlist_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        }
                        if (chbox_AutoFixCutPartLength.Checked)
                            resizeView.AutoResizeView(view, offsetFromFrame, minCutpartDecrease);
                        if (chbox_AutoCreateSection.Checked)
                        {
                            tsd.View.ViewAttributes viewAttributes = new tsd.View.ViewAttributes();
                            viewAttributes.LoadAttributes(cb_SecAttribute.Text);
                            tsd.SectionMarkBase.SectionMarkAttributes sectionMarkAttributes = new SectionMarkBase.SectionMarkAttributes();
                            sectionMarkAttributes.LoadAttributes(cb_SecMarkAttribute.Text);
                            tsd.View secView = null;
                            tsd.SectionMark secMark = null;
                            double purlinLength = function.GetLength(modelPart);
                            tsd.View.CreateSectionView(view, pointXminYmax, pointXminYmin, pointXminYmin, purlinLength + 50, 40, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                            if (secView.Attributes.Scale > view.Attributes.Scale)
                            {
                                secView.Attributes.Scale = view.Attributes.Scale;
                                secView.Modify();
                            }
                            secView.Origin = new t3d.Point(view.Origin.X + secView.Width, view.Origin.Y - (secView.Height + view.Height * 0.5));
                            secView.Modify();
                            CreateAutoDimSection(secView);
                        }
                    }
                    if (chbox_AddPartMark.Checked)
                        Add_Mark();
                }
                else if (drobj is tsd.View)
                {
                    tsd.View view = drobj as tsd.View;
                    if (chbox_ClearCurrentDim.Checked)
                        clearDrawingObjects.ClearDim(view);
                    double viewScale = view.Attributes.Scale;
                    if (dimSpace == 0 && dimSpaceDim2Dim == 0)
                    {
                        dimSpace = Convert.ToInt32(viewScale * 3 * 5);
                        dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
                    }
                    tsm.TransformationPlane viewplane = new TransformationPlane(view.DisplayCoordinateSystem);
                    wph.SetCurrentTransformationPlane(viewplane);
                    tsd.DrawingObjectEnumerator drawingPartEnumerator = view.GetAllObjects(new Type[] { typeof(tsd.Part) });
                    //khai báo viewbase
                    tsd.ViewBase viewBase = view as tsd.ViewBase;
                    foreach (tsd.DrawingObject part in drawingPartEnumerator)
                    {
                        //Convert về Part
                        tsd.Part drpart = part as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                        string partProfileType = function.GetProfileType(mpart);
                        //tính toán điểm của part này trên view bản vẽ
                        Part_Edge partEdge = new Part_Edge(view, mpart); //tạo 1 myPartEdge thuộc kiểu dữ liệu (class) Part_Edge
                                                                         //model.CommitChanges();
                        List<t3d.Point> pointListPart = partEdge.List_Edge;
                        ArrayList listPointCenter = mpart.GetCenterLine(true);
                        t3d.Point clPoint1 = listPointCenter[0] as t3d.Point;
                        t3d.Point clPoint2 = listPointCenter[1] as t3d.Point;
                        t3d.Point mid_YofPart = partEdge.PointMidLeft;
                        t3d.Point pointXmaxYmax = partEdge.PointXmaxYmax;
                        t3d.Point pointXmaxYmin = partEdge.PointXmaxYmin;
                        t3d.Point pointXminYmax = partEdge.PointXminYmax;
                        t3d.Point pointXminYmin = partEdge.PointXminYmin;

                        t3d.Point pointXmin0 = partEdge.PointXmin0;
                        t3d.Point pointXmin1 = partEdge.PointXmin1;
                        t3d.Point pointXmax0 = partEdge.PointXmax0;
                        t3d.Point pointXmax1 = partEdge.PointXmax1;
                        t3d.Point pointYmin0 = partEdge.PointYmin0;
                        t3d.Point pointYmin1 = partEdge.PointYmin1;
                        t3d.Point pointYmax0 = partEdge.PointYmax0;
                        t3d.Point pointYmax1 = partEdge.PointYmax1;

                        dimOverall_X.Add(pointXminYmax);
                        dimOverall_X.Add(pointXmaxYmax);
                        dimOverall_Y.Add(pointXminYmax);
                        dimOverall_Y.Add(pointXminYmin);

                        Boltlist_Bolt_XY.Add(pointXminYmax);
                        Boltlist_Bolt_XY.Add(pointXmaxYmax);
                        Boltlist_Bolt_XY.Add(pointXminYmin);

                        Boltlist_Bolt_Y_top.Add(pointXminYmax);
                        Boltlist_Bolt_Y_top.Add(pointXmaxYmax);
                        Boltlist_Bolt_Y_bot.Add(pointXminYmin);
                        Boltlist_Bolt_Y_bot.Add(pointXmaxYmin);
                        tsm.ModelObjectEnumerator boltenum = mpart.GetBolts();
                        foreach (tsm.BoltGroup bgr in boltenum)
                        {
                            Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                            //Tìm bolt theo phương XY
                            tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY;
                            tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                            //wph.SetCurrentTransformationPlane(viewplane);
                            if (bolt_xy != null)
                            {
                                bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                foreach (t3d.Point p in bolt_xy.BoltPositions)
                                {
                                    Boltlist_Bolt_XY.Add(p);
                                }
                            }
                            if (bolt_y != null)
                            {
                                bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                foreach (t3d.Point p in bolt_y.BoltPositions)
                                {
                                    if (p.Y > mid_YofPart.Y)
                                    {
                                        Boltlist_Bolt_Y_top.Add(p);
                                    }
                                    else
                                    {
                                        Boltlist_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                        }
                        boltenum.Reset();
                        //Tất cả những liên quan đến khởi tạo dim sẽ có câu Handler
                        tsd.StraightDimensionSetHandler smyDrawingHandler = new StraightDimensionSetHandler();
                        //Nếu 2 điểm center line trùng nhau thì là dim section
                        if (Math.Abs(clPoint1.X - clPoint2.X) < 3)
                        {
                            //Trường hợp Xà gồ Z móc trên quay qua bên Trái
                            if (pointXmin1 == pointYmax0 && partProfileType == "Z")
                            {
                                // MessageBox.Show("Quay trai");
                                dimOverall_X.Clear();
                                dimOverall_X_.Clear();
                                dimOverall_Y.Clear();
                                dimOverall_X.Add(pointXmin1);
                                dimOverall_X.Add(pointYmax1);
                                dimOverall_X_.Add(pointYmin0);
                                dimOverall_X_.Add(pointYmin1);
                                dimOverall_Y.Add(pointXmin1);
                                dimOverall_Y.Add(pointYmin0);
                                dimOverall_Y1.Add(pointXmin1);
                                dimOverall_Y1.Add(pointXmin0);
                                dimOverall_Y2.Add(pointXmax1);
                                dimOverall_Y2.Add(pointXmax0);
                                //tạo dim X lên trên
                                createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                //tạo dim X xuống dưới
                                createDimension.CreateStraightDimensionSet_X_(viewBase, dimOverall_X_, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                //tạo dim Y qua trái
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                //tạo dim Y1 móc trên
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y1, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                                //tạo dim Y1 móc trên
                                createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y2, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                                dim_Bolt_Y_top.Add(pointXmin1);
                                dim_Bolt_Y_top.Add(pointYmax1);
                                dim_Bolt_Y_bot.Add(pointYmin0);
                                dim_Bolt_Y_bot.Add(pointYmin1);
                                dim_Bolt_X.Add(pointXmin1);
                                dim_Bolt_X.Add(pointYmin0);
                                foreach (tsm.BoltGroup bgr in boltenum)
                                {
                                    try
                                    {
                                        Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                        //Tìm bolt theo phương XY
                                        tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                        tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                        //wph.SetCurrentTransformationPlane(viewplane);
                                        if (bolt_x != null)
                                        {
                                            bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_x.BoltPositions)
                                                dim_Bolt_X.Add(p);
                                        }
                                        if (bolt_y != null)
                                        {
                                            bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_y.BoltPositions)
                                            {
                                                if (p.Y > mid_YofPart.Y)
                                                    dim_Bolt_Y_top.Add(p);
                                                else
                                                    dim_Bolt_Y_bot.Add(p);
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                                if (dim_Bolt_Y_top.Count != 2)
                                    createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                                if (dim_Bolt_Y_bot.Count != 2)
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                                if (dim_Bolt_X.Count != 2)
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            }
                            //Trường hợp Xà gồ Z móc trên quay qua bên Phải
                            else if (pointXmax1 == pointYmax1 && partProfileType == "Z")
                            {
                                dimOverall_X.Clear();
                                dimOverall_X_.Clear();
                                dimOverall_Y.Clear();
                                dimOverall_X.Add(pointYmax0);
                                dimOverall_X.Add(pointYmax1);
                                dimOverall_X_.Add(pointXmin0);
                                dimOverall_X_.Add(pointYmin1);
                                dimOverall_Y.Add(pointYmax1);
                                dimOverall_Y.Add(pointYmin1);
                                dimOverall_Y1.Add(pointXmax1);
                                dimOverall_Y1.Add(pointXmax0);
                                dimOverall_Y2.Add(pointXmin1);
                                dimOverall_Y2.Add(pointXmin0);
                                //tạo dim X lên trên
                                createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                //tạo dim X xuống dưới
                                createDimension.CreateStraightDimensionSet_X_(viewBase, dimOverall_X_, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                //tạo dim Y qua trái
                                createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                //tạo dim Y1 móc trên
                                createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y1, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                                //tạo dim Y1 móc trên
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y2, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                                dim_Bolt_Y_top.Add(pointYmax0);
                                dim_Bolt_Y_top.Add(pointYmax1);
                                dim_Bolt_Y_bot.Add(pointXmin0);
                                dim_Bolt_Y_bot.Add(pointYmin1);
                                dim_Bolt_X.Add(pointYmax1);
                                dim_Bolt_X.Add(pointYmin1);
                                foreach (tsm.BoltGroup bgr in boltenum)
                                {
                                    Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                    //Tìm bolt theo phương XY
                                    tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                    tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                    //wph.SetCurrentTransformationPlane(viewplane);
                                    if (bolt_x != null)
                                    {
                                        bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                        foreach (t3d.Point p in bolt_x.BoltPositions)
                                            dim_Bolt_X.Add(p);
                                    }
                                    if (bolt_y != null)
                                    {
                                        bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                        foreach (t3d.Point p in bolt_y.BoltPositions)
                                        {
                                            if (p.Y > mid_YofPart.Y)
                                                dim_Bolt_Y_top.Add(p);
                                            else
                                                dim_Bolt_Y_bot.Add(p);
                                        }
                                    }
                                }
                                if (dim_Bolt_Y_top.Count != 2)
                                    createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                                if (dim_Bolt_Y_bot.Count != 2)
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                                if (dim_Bolt_X.Count != 2)
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            }
                            //Trường hợp xà gồ C
                            else if (partProfileType == "C")
                            {
                                //tạo ds các điểm có cùng tọa độ X min, Nếu số phần từ là 4 thì C quay lưng về bên trái và ngược lại.
                                List<t3d.Point> pointListPart_minX = pointListPart.FindAll(p => Math.Abs(p.X - pointXmin0.X) < 0.1).ToList();
                                //C quay lung qua phai

                                dimOverall_X.Add(pointYmax0);
                                dimOverall_X.Add(pointYmax1);
                                //Nếu C quay lưng qua trái
                                if (pointListPart_minX.Count == 4)
                                {
                                    dimOverall_Y.Add(pointYmax1);
                                    dimOverall_Y.Add(pointYmin1);
                                }
                                //Nếu C quay lưng qua phải
                                else
                                {
                                    dimOverall_Y.Add(pointYmax0);
                                    dimOverall_Y.Add(pointYmin0);
                                }
                                dim_Bolt_Y_top.Add(pointYmax0);
                                dim_Bolt_Y_top.Add(pointYmax1);
                                dim_Bolt_Y_bot.Add(pointYmin0);
                                dim_Bolt_Y_bot.Add(pointYmin1);
                                //Nếu C quay lưng qua trái
                                if (pointListPart_minX.Count == 4)
                                {
                                    dim_Bolt_X.Add(pointYmax1);
                                    dim_Bolt_X.Add(pointYmin1);
                                }
                                //Nếu C quay lưng qua phải
                                else
                                {
                                    dim_Bolt_X.Add(pointYmax0);
                                    dim_Bolt_X.Add(pointYmin0);
                                }
                                //tạo dim X lên trên
                                createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                //tạo dim Y qua trái
                                if (pointListPart_minX.Count == 4)
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                else //tạo dim Y qua phải
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                                foreach (tsm.BoltGroup bgr in boltenum)
                                {
                                    Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                    //Tìm bolt theo phương XY
                                    tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                    tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                    //wph.SetCurrentTransformationPlane(viewplane);
                                    if (bolt_x != null)
                                    {
                                        bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                        foreach (t3d.Point p in bolt_x.BoltPositions)
                                            dim_Bolt_X.Add(p);
                                    }
                                    if (bolt_y != null)
                                    {
                                        bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                        foreach (t3d.Point p in bolt_y.BoltPositions)
                                        {
                                            if (p.Y > mid_YofPart.Y)
                                                dim_Bolt_Y_top.Add(p);
                                            else
                                                dim_Bolt_Y_bot.Add(p);
                                        }
                                    }
                                }
                                if (dim_Bolt_Y_top.Count != 2)
                                    createDimension.CreateStraightDimensionSet_X(viewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                                if (dim_Bolt_Y_bot.Count != 2)
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim);

                                if (dim_Bolt_X.Count != 2)
                                {
                                    if (pointListPart_minX.Count == 4)
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                    else //tạo dim Y qua phải
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                }
                            }
                        }
                        //Nếu không sẽ dim purlin bình thường
                        else
                        {
                            //Tạo dimension bao theo phương X
                            createDimension.CreateStraightDimensionSet_X(viewBase, dimOverall_X, dimAttribute, dimSpace);
                            //Tạo dimension bao theo phương Y
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, dimOverall_Y, dimAttribute, dimSpace);
                            //Tạo bolt xy theo phương X
                            createDimension.CreateStraightDimensionSet_X(viewBase, Boltlist_Bolt_XY, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //Tạo bolt xy theo phương Y
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, Boltlist_Bolt_XY, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //Tạo bolt Y theo phương X
                            if (Boltlist_Bolt_Y_top.Count > 2)
                            {
                                createDimension.CreateStraightDimensionSet_X(viewBase, Boltlist_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            }
                            if (Boltlist_Bolt_Y_bot.Count > 2)
                            {
                                createDimension.CreateStraightDimensionSet_X_(viewBase, Boltlist_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            }
                            if (chbox_AutoFixCutPartLength.Checked)
                                resizeView.AutoResizeView(view, offsetFromFrame, minCutpartDecrease);
                            if (chbox_AutoCreateSection.Checked)
                            {
                                tsd.View.ViewAttributes viewAttributes = new tsd.View.ViewAttributes();
                                viewAttributes.LoadAttributes(cb_SecAttribute.Text);
                                tsd.SectionMarkBase.SectionMarkAttributes sectionMarkAttributes = new SectionMarkBase.SectionMarkAttributes();
                                sectionMarkAttributes.LoadAttributes(cb_SecMarkAttribute.Text);
                                tsd.View secView = null;
                                tsd.SectionMark secMark = null;
                                double purlinLength = function.GetLength(mpart);
                                tsd.View.CreateSectionView(view, pointXminYmax, pointXminYmin, pointXminYmin, purlinLength + 50, 40, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                                if (secView.Attributes.Scale > view.Attributes.Scale)
                                {
                                    secView.Attributes.Scale = view.Attributes.Scale;
                                    secView.Modify();
                                }
                                secView.Origin = new t3d.Point(view.Origin.X + secView.Width, view.Origin.Y - (secView.Height + view.Height * 0.5));
                                secView.Modify();
                                CreateAutoDimSection(secView);
                            }
                        }
                        if (chbox_AddPartMark.Checked)
                            Add_Mark();
                    }
                }
            }
        }

        public void CreateAutoDimSection(tsd.View view)
        {
            CreateDimension createDimension = new CreateDimension();
            ClearDrawingObjects clearDrObjs = new ClearDrawingObjects();
            Function function = new Function();
            if (chbox_ClearCurrentDim.Checked)
                clearDrObjs.ClearPartMark(view);

            string dimAttribute = cb_DimentionAttributes.Text; //Khoảng cb_DimentionAttributes mặc định
            int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text); //Khoảng dimSpace 1
            int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
            double viewScale = view.Attributes.Scale;
            if (dimSpace == 0 && dimSpaceDim2Dim == 0)
            {
                dimSpace = Convert.ToInt32(viewScale * 3 * 5);
                dimSpaceDim2Dim = Convert.ToInt32(viewScale * 5);
            }
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler(); //Lay workplanhandler
                                                                    //Tạo tọa độ làm việc gốc (trả về gốc)
            tsm.TransformationPlane OriginalModelPlane = new TransformationPlane();
            wph.SetCurrentTransformationPlane(OriginalModelPlane);

            tsd.PointList dim_Bolt_Y_top = new PointList(); //ds các điểm bolt Y nằm trên center line
            tsd.PointList dim_Bolt_Y_bot = new PointList(); //ds các điểm bolt Y nằm dưới center line
            tsd.PointList dim_Bolt_X = new PointList();
            tsd.PointList dimOverall_X = new PointList();
            tsd.PointList dimOverall_X_ = new PointList();
            tsd.PointList dimOverall_Y = new PointList();
            tsd.PointList dimOverall_Y1 = new PointList();
            tsd.PointList dimOverall_Y2 = new PointList();

            //khai báo viewbase
            tsd.ViewBase MyViewBase = view as tsd.ViewBase;
            // Lấy hệ tọa độ của bản vẽ.
            tsm.TransformationPlane viewplane = new TransformationPlane(view.DisplayCoordinateSystem);
            //Chuyển hệ tọa độ làm việc về hệ tọa độ của view
            wph.SetCurrentTransformationPlane(viewplane);
            //lấy các đối tượng trong view.
            tsd.DrawingObjectEnumerator drobjenum = view.GetModelObjects();
            try
            {
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    //kiểm tra xem đối tượng này có phải là part không ?
                    if (drobj is tsd.Part)
                    {
                        //Convert đối tượng về part
                        tsd.Part dr_part = drobj as tsd.Part;
                        //chuyển dr_part về model
                        tsm.Part model_part = model.SelectModelObject(dr_part.ModelIdentifier) as tsm.Part;
                        string partProfileType = string.Empty;
                        model_part.GetReportProperty("PROFILE_TYPE", ref partProfileType);
                        //tính toán điểm của part này trên view bản vẽ
                        Part_Edge partEdge = new Part_Edge(view, model_part); //tạo 1 myPartEdge thuộc kiểu dữ liệu (class) Part_Edge
                        List<t3d.Point> pointListPart = partEdge.List_Edge;
                        t3d.Point mid_YofPart = partEdge.PointMidLeft;
                        t3d.Point pointXmin0 = partEdge.PointXmin0;
                        t3d.Point pointXmin1 = partEdge.PointXmin1;
                        t3d.Point pointXmax0 = partEdge.PointXmax0;
                        t3d.Point pointXmax1 = partEdge.PointXmax1;
                        t3d.Point pointYmin0 = partEdge.PointYmin0;
                        t3d.Point pointYmin1 = partEdge.PointYmin1;
                        t3d.Point pointYmax0 = partEdge.PointYmax0;
                        t3d.Point pointYmax1 = partEdge.PointYmax1;
                        tsm.ModelObjectEnumerator boltenum = model_part.GetBolts();

                        //Trường hợp Xà gồ Z móc trên quay qua bên Trái
                        if (pointXmin1 == pointYmax0 && partProfileType == "Z")
                        {
                            dimOverall_X.Add(pointXmin1);
                            dimOverall_X.Add(pointYmax1);
                            dimOverall_X_.Add(pointYmin0);
                            dimOverall_X_.Add(pointYmin1);
                            dimOverall_Y.Add(pointXmin1);
                            dimOverall_Y.Add(pointYmin0);
                            dimOverall_Y1.Add(pointXmin1);
                            dimOverall_Y1.Add(pointXmin0);
                            dimOverall_Y2.Add(pointXmax1);
                            dimOverall_Y2.Add(pointXmax0);
                            //tạo dim X lên trên
                            createDimension.CreateStraightDimensionSet_X(MyViewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim X xuống dưới
                            createDimension.CreateStraightDimensionSet_X_(MyViewBase, dimOverall_X_, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim Y qua trái
                            createDimension.CreateStraightDimensionSet_Y_(MyViewBase, dimOverall_Y, dimAttribute, dimSpace);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y_(MyViewBase, dimOverall_Y1, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y(MyViewBase, dimOverall_Y2, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            dim_Bolt_Y_top.Add(pointXmin1);
                            dim_Bolt_Y_top.Add(pointYmax1);
                            dim_Bolt_Y_bot.Add(pointYmin0);
                            dim_Bolt_Y_bot.Add(pointYmin1);
                            dim_Bolt_X.Add(pointXmin1);
                            dim_Bolt_X.Add(pointYmin0);
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                //wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_x != null)
                                {
                                    bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_x.BoltPositions)
                                        dim_Bolt_X.Add(p);
                                }
                                if (bolt_y != null)
                                {
                                    bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_y.BoltPositions)
                                    {
                                        if (p.Y > mid_YofPart.Y)
                                            dim_Bolt_Y_top.Add(p);
                                        else
                                            dim_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                            if (dim_Bolt_Y_top.Count != 2)
                                createDimension.CreateStraightDimensionSet_X(MyViewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_Y_bot.Count != 2)
                                createDimension.CreateStraightDimensionSet_X_(MyViewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_X.Count != 2)
                                createDimension.CreateStraightDimensionSet_Y_(MyViewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        }
                        //Trường hợp Xà gồ Z móc trên quay qua bên Phải
                        else if (pointXmax1 == pointYmax1 && partProfileType == "Z")
                        {
                            dimOverall_X.Add(pointYmax0);
                            dimOverall_X.Add(pointYmax1);
                            dimOverall_X_.Add(pointXmin0);
                            dimOverall_X_.Add(pointYmin1);
                            dimOverall_Y.Add(pointYmax1);
                            dimOverall_Y.Add(pointYmin1);
                            dimOverall_Y1.Add(pointXmax1);
                            dimOverall_Y1.Add(pointXmax0);
                            dimOverall_Y2.Add(pointXmin1);
                            dimOverall_Y2.Add(pointXmin0);
                            //tạo dim X lên trên
                            createDimension.CreateStraightDimensionSet_X(MyViewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim X xuống dưới
                            createDimension.CreateStraightDimensionSet_X_(MyViewBase, dimOverall_X_, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim Y qua trái
                            createDimension.CreateStraightDimensionSet_Y(MyViewBase, dimOverall_Y, dimAttribute, dimSpace);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y(MyViewBase, dimOverall_Y1, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            //tạo dim Y1 móc trên
                            createDimension.CreateStraightDimensionSet_Y_(MyViewBase, dimOverall_Y2, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);
                            dim_Bolt_Y_top.Add(pointYmax0);
                            dim_Bolt_Y_top.Add(pointYmax1);
                            dim_Bolt_Y_bot.Add(pointXmin0);
                            dim_Bolt_Y_bot.Add(pointYmin1);
                            dim_Bolt_X.Add(pointYmax1);
                            dim_Bolt_X.Add(pointYmin1);
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                //wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_x != null)
                                {
                                    bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_x.BoltPositions)
                                        dim_Bolt_X.Add(p);
                                }
                                if (bolt_y != null)
                                {
                                    bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_y.BoltPositions)
                                    {
                                        if (p.Y > mid_YofPart.Y)
                                            dim_Bolt_Y_top.Add(p);
                                        else
                                            dim_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                            if (dim_Bolt_Y_top.Count != 2)
                                createDimension.CreateStraightDimensionSet_X(MyViewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_Y_bot.Count != 2)
                                createDimension.CreateStraightDimensionSet_X_(MyViewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_X.Count != 2)
                                createDimension.CreateStraightDimensionSet_Y(MyViewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                        }
                        //Trường hợp xà gồ C
                        else if (partProfileType == "C")
                        {
                            //tạo ds các điểm có cùng tọa độ X min, Nếu số phần từ là 4 thì C quay lưng về bên trái và ngược lại.
                            List<t3d.Point> pointListPart_minX = pointListPart.FindAll(p => Math.Abs(p.X - pointXmin0.X) < 0.1).ToList();
                            //C quay lung qua phai

                            dimOverall_X.Add(pointYmax0);
                            dimOverall_X.Add(pointYmax1);
                            //Nếu C quay lưng qua trái
                            if (pointListPart_minX.Count == 4)
                            {
                                dimOverall_Y.Add(pointYmax1);
                                dimOverall_Y.Add(pointYmin1);
                            }
                            //Nếu C quay lưng qua phải
                            else
                            {
                                dimOverall_Y.Add(pointYmax0);
                                dimOverall_Y.Add(pointYmin0);
                            }
                            dim_Bolt_Y_top.Add(pointYmax0);
                            dim_Bolt_Y_top.Add(pointYmax1);
                            dim_Bolt_Y_bot.Add(pointYmin0);
                            dim_Bolt_Y_bot.Add(pointYmin1);
                            //Nếu C quay lưng qua trái
                            if (pointListPart_minX.Count == 4)
                            {
                                dim_Bolt_X.Add(pointYmax1);
                                dim_Bolt_X.Add(pointYmin1);
                            }
                            //Nếu C quay lưng qua phải
                            else
                            {
                                dim_Bolt_X.Add(pointYmax0);
                                dim_Bolt_X.Add(pointYmin0);
                            }
                            //tạo dim X lên trên
                            createDimension.CreateStraightDimensionSet_X(MyViewBase, dimOverall_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            //tạo dim Y qua trái
                            if (pointListPart_minX.Count == 4)
                                createDimension.CreateStraightDimensionSet_Y(MyViewBase, dimOverall_Y, dimAttribute, dimSpace);
                            else //tạo dim Y qua phải
                                createDimension.CreateStraightDimensionSet_Y_(MyViewBase, dimOverall_Y, dimAttribute, dimSpace);
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(view, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_x = tinhbolt.Bolt_X;
                                tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                //wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_x != null)
                                {
                                    bolt_x.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_x.BoltPositions)
                                        dim_Bolt_X.Add(p);
                                }
                                if (bolt_y != null)
                                {
                                    bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_y.BoltPositions)
                                    {
                                        if (p.Y > mid_YofPart.Y)
                                            dim_Bolt_Y_top.Add(p);
                                        else
                                            dim_Bolt_Y_bot.Add(p);
                                    }
                                }
                            }
                            if (dim_Bolt_Y_top.Count != 2)
                                createDimension.CreateStraightDimensionSet_X(MyViewBase, dim_Bolt_Y_top, dimAttribute, dimSpace - 2 * dimSpaceDim2Dim);

                            if (dim_Bolt_Y_bot.Count != 2)
                                createDimension.CreateStraightDimensionSet_X_(MyViewBase, dim_Bolt_Y_bot, dimAttribute, dimSpace - dimSpaceDim2Dim);

                            if (dim_Bolt_X.Count != 2)
                            {
                                if (pointListPart_minX.Count == 4)
                                    createDimension.CreateStraightDimensionSet_Y(MyViewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                                else //tạo dim Y qua phải
                                    createDimension.CreateStraightDimensionSet_Y_(MyViewBase, dim_Bolt_X, dimAttribute, dimSpace - dimSpaceDim2Dim);
                            }
                        }
                    }
                }
            }
            catch { }
            wph.SetCurrentTransformationPlane(OriginalModelPlane);
        } //Tạo dim cho section chỉ có 1 part 

        public void DimArrestor()
        {
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            wph.SetCurrentTransformationPlane(new TransformationPlane());
            CreateDimension createDimension = new CreateDimension();
            ClearDrawingObjects clearDrObjs = new ClearDrawingObjects();
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            string dimAttributes = cb_DimentionAttributes.Text; //Khoảng cb_DimentionAttributes mặc định
            int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
            int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);

            wph.SetCurrentTransformationPlane(new TransformationPlane());

            foreach (tsd.DrawingObject drobj in drobjenum)
            {
                if (drobj is tsd.Part)
                {
                    //Convert đối tượng về part
                    tsd.Part drPart = drobj as tsd.Part;
                    //chuyển dr_part về model
                    tsm.Part mPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                    tsd.View view = drPart.GetView() as tsd.View;
                    //khai báo viewbase
                    tsd.ViewBase viewBase = view as tsd.ViewBase;
                    if (chbox_ClearCurrentDim.Checked)
                        clearDrObjs.ClearDim(view);
                    wph.SetCurrentTransformationPlane(new TransformationPlane(view.DisplayCoordinateSystem));
                    ArrayList arlistRefLinePoint = mPart.GetReferenceLine(true);
                    //MessageBox.Show(arlistRefLinePoint.Count.ToString());
                    t3d.Point p1 = arlistRefLinePoint[0] as t3d.Point;
                    t3d.Point p2 = arlistRefLinePoint[1] as t3d.Point;
                    t3d.Point p3 = arlistRefLinePoint[2] as t3d.Point;
                    t3d.Point p4 = arlistRefLinePoint[3] as t3d.Point;
                    PointList listPointsTotal_X = new PointList { p1, p4 }; //ds điểm p1 p2
                    PointList listPointsTotal_X1234 = new PointList { p1, p2, p3, p4 }; //ds điểm p1 p2
                    PointList listPointsTotal_Y12 = new PointList { p1, p2 };
                    PointList listPointsTotal_Y34 = new PointList { p3, p4 };
                    createDimension.CreateStraightDimensionSet_FX(p3, p2, viewBase, listPointsTotal_X, dimAttributes, dimSpace);
                    createDimension.CreateStraightDimensionSet_FX(p3, p2, viewBase, listPointsTotal_X1234, dimAttributes, dimSpace - dimSpaceDim2Dim);
                    createDimension.CreateStraightDimensionSet_FY_(p2, p3, viewBase, listPointsTotal_Y12, dimAttributes, dimSpace - dimSpaceDim2Dim);
                    createDimension.CreateStraightDimensionSet_FX_(p1, p2, viewBase, listPointsTotal_Y12, dimAttributes, dimSpace - dimSpaceDim2Dim);
                    createDimension.CreateStraightDimensionSet_FY_(p2, p3, viewBase, listPointsTotal_Y34, dimAttributes, dimSpace - dimSpaceDim2Dim);
                    createDimension.CreateStraightDimensionSet_FX(p3, p4, viewBase, listPointsTotal_Y34, dimAttributes, dimSpace - dimSpaceDim2Dim);
                    //createDimension.CreateStraightDimensionSet_X(viewBase, listPointsTotal_X, dimAttributes, dimSpace + dimSpaceDim2Dim);
                    //createDimension.CreateStraightDimensionSet_X(viewBase, listPointsTotal_X1234, dimAttributes, dimSpace - dimSpaceDim2Dim);
                    //createDimension.CreateStraightDimensionSet_Y(viewBase, listPointsTotal_Y12, dimAttributes, dimSpace - dimSpaceDim2Dim);
                }
            }
            drawingHandler.SaveActiveDrawing();
        }

        private void btn_DimGA_X_Click(object sender, EventArgs e) // DIM THEO PHUONG X
        {
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
            CreateDimension createDimension = new CreateDimension(); // tạo 1 createDimension xem thêm class CreateDimension
            wph.SetCurrentTransformationPlane(new TransformationPlane());
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            tsm.Part mainPart = null;
            PartDistribution partDistribution = null;
            if (actived_dr is tsd.AssemblyDrawing)
            {
                tsd.AssemblyDrawing assemblyDrawing = actived_dr as tsd.AssemblyDrawing; //Lấy bản vẽ assembly
                tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                mainPart = assembly.GetMainPart() as tsm.Part;
                partDistribution = new PartDistribution(mainPart, drobjenum);
            }
            // Lấy đối tượng được chọn trên bản vẽ
            tsdui.Picker c_picker = drawingHandler.GetPicker();
            var pick = c_picker.PickPoint("Pick first dimension point");
            t3d.Point point2 = new t3d.Point();
            //try
            //{
            //var pick2 = c_picker.PickPoint("Pick second dimension point");
            //point2 = pick2.Item1;
            //}
            //catch { }

            var pick3 = c_picker.PickPoint("Pick location for dimension");
            t3d.Point point3 = pick3.Item1;

            //Lấy tọa độ điểm pick (điểm gốc)
            t3d.Point point1 = pick.Item1;
            //Lấy view hiện hành thông qua điểm vừa pick
            tsd.View viewcur = pick.Item2 as tsd.View;
            tsd.ViewBase viewbase = pick.Item2;
            //tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
            //wph.SetCurrentTransformationPlane(viewplane);
            //model.CommitChanges();

            tsd.PointList pointListDimRef = new tsd.PointList() { point1 };
            tsd.PointList pointListDimCenter = new tsd.PointList() { point1 };

            tsd.PointList pointListDimEde_X = new PointList() { point1 };//Tập hợp điểm X bên trên
            tsd.PointList pointListDimEde_X_ = new PointList() { point1 };//Tập hợp điểm X bên dưới
            tsd.PointList pointListDimEde_Y = new PointList() { point1 };//Tập hợp điểm Y bên phải
            tsd.PointList pointListDimEde_Y_ = new PointList() { point1 }; //Tập hợp điểm Y bên trái

            //if (point2.X != 0 && point2.Y != 0)
            //{
            //    pointListDimRef.Add(point2);
            //    pointListDimCenter.Add(point2);
            //    pointListDimEde_X.Add(point2);
            //    pointListDimEde_X_.Add(point2);
            //    pointListDimEde_Y.Add(point2);
            //    pointListDimEde_Y_.Add(point2);
            //}

            foreach (tsd.DrawingObject drobj in drobjenum)
            {
                if (drobj is tsd.Part)
                {
                    //Convert về Part
                    tsd.Part drPart = drobj as tsd.Part;
                    //Lấy model part thông qua part model identifỉe
                    tsm.Part mPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                    Part_Edge part_edge = new Part_Edge(viewcur, mPart);
                    t3d.Point maxX = part_edge.PointXmax0;
                    t3d.Point minX = part_edge.PointXmin0;
                    t3d.Point maxY = part_edge.PointYmax0;
                    t3d.Point minY = part_edge.PointYmin0;
                    pointListDimEde_X.Add(part_edge.PointXminYmax);
                    pointListDimEde_X.Add(part_edge.PointXmaxYmax);
                    pointListDimEde_X_.Add(part_edge.PointXminYmin);
                    pointListDimEde_X_.Add(part_edge.PointXmaxYmin);
                    pointListDimEde_Y.Add(part_edge.PointXmaxYmax);
                    pointListDimEde_Y.Add(part_edge.PointXmaxYmin);
                    pointListDimEde_Y_.Add(part_edge.PointXminYmax);
                    pointListDimEde_Y_.Add(part_edge.PointXminYmin);
                    ArrayList listRefLine = mPart.GetReferenceLine(false);
                    ArrayList listCenterLine = mPart.GetCenterLine(true);
                    foreach (t3d.Point p in listRefLine)
                    {
                        pointListDimRef.Add(p);
                    }

                    foreach (t3d.Point p in listCenterLine)
                    {
                        pointListDimCenter.Add(p);
                    }
                }
            }
            drobjenum.Reset();

            if (imcb_SelectDimRefCenEdge.SelectedIndex == 1) //NEU CHON DIM REFERENCE LINE
            {
                int spaceDim = Convert.ToInt32(Math.Abs(pointListDimRef[0].Y - point3.Y)); //Tinh khoang cach tu diem dim dau tien, đến điểm đặt dim
                if (point1.Y < point3.Y)
                    createDimension.CreateStraightDimensionSet_X(viewbase, pointListDimRef, cb_DimentionAttributes.Text, spaceDim);
                else
                    createDimension.CreateStraightDimensionSet_X_(viewbase, pointListDimRef, cb_DimentionAttributes.Text, spaceDim);
            } //NEU CHON DIM REFERENCE LINE
            else if (imcb_SelectDimRefCenEdge.SelectedIndex == 2) //NEU CHON DIM CENTER LINE
            {
                int spaceDim = Convert.ToInt32(Math.Abs(pointListDimCenter[0].Y - point3.Y)); //Tinh khoang cach tu diem dim dau tien, đến điểm đặt dim
                if (point1.Y < point3.Y)
                {
                    createDimension.CreateStraightDimensionSet_X(viewbase, pointListDimCenter, cb_DimentionAttributes.Text, spaceDim);
                }
                else
                {
                    createDimension.CreateStraightDimensionSet_X_(viewbase, pointListDimCenter, cb_DimentionAttributes.Text, spaceDim);
                }
            } //NEU CHON DIM CENTER LINE
            else if (imcb_SelectDimRefCenEdge.SelectedIndex == 3) //NEU CHON DIM EDGE
            {
                int spaceDim = Convert.ToInt32(Math.Abs(pointListDimEde_X[0].Y - point3.Y)); //Tinh khoang cach tu diem dim dau tien, đến điểm đặt dim
                if (point1.Y < point3.Y)
                {
                    createDimension.CreateStraightDimensionSet_X(viewbase, pointListDimEde_X, cb_DimentionAttributes.Text, spaceDim);
                }
                else
                {
                    createDimension.CreateStraightDimensionSet_X_(viewbase, pointListDimEde_X_, cb_DimentionAttributes.Text, spaceDim);
                }
            } //NEU CHON DIM EDGE
            else if (imcb_SelectDimRefCenEdge.SelectedIndex == 0 || imcb_SelectDimRefCenEdge.SelectedIndex == -1) //NEU CHON DIM PARTS OR BOLTS ASSEMBLY
            {
                if (partDistribution.havePart == true)
                {
                    List<tsm.Part> platesTamThayBoltTron = partDistribution.ListTamThayBoltTron;
                    List<tsm.Part> platesTamDung = partDistribution.ListTamDung;
                    List<tsm.Part> platesTamNgang = partDistribution.ListTamNgang;
                    List<tsm.Part> platesNghiengPhai = partDistribution.ListTamNghiengPhai;
                    List<tsm.Part> platesNghiengTrai = partDistribution.ListTamNghiengTrai;

                    List<tsm.Part> partsIDung = partDistribution.ListPartIDung;
                    List<tsm.Part> partsINgang = partDistribution.ListPartINgang;
                    List<tsm.Part> partsINghiengTrai = partDistribution.ListPartINghiengTrai;
                    List<tsm.Part> partsINghiengPhai = partDistribution.ListPartINghiengPhai;
                    List<tsm.Part> partsLUsection = partDistribution.ListPartLUsection;

                    tsd.PointList plpartLeftDimX = new PointList(); //Tập hợp điểm DIM mép trái PL phuong X
                    tsd.PointList plpartRightDimX = new PointList(); //Tập hợp điểm DIM mép phai PL phuong X
                    tsd.PointList plpartLeftDimX_ = new PointList();//Tập hợp điểm DIM mép trái PL phuong X_
                    tsd.PointList plpartRightDimX_ = new PointList();//Tập hợp điểm DIM mép phai PL phuong X_

                    tsd.PointList plpartTopDimY = new PointList(); //Tập hợp điểm DIM mép tren PL phuong Y
                    tsd.PointList plpartBotDimY = new PointList(); //Tập hợp điểm DIM mép duoi PL phuong Y
                    tsd.PointList plpartTopDimY_ = new PointList();//Tập hợp điểm DIM mép tren PL phuong Y_
                    tsd.PointList plpartBotDimY_ = new PointList();//Tập hợp điểm DIM mép duoi PL phuong Y_

                    plpartLeftDimX.Add(point1);
                    plpartRightDimX.Add(point1);
                    plpartLeftDimX_.Add(point1);
                    plpartRightDimX_.Add(point1);

                    plpartTopDimY.Add(point1);
                    plpartBotDimY.Add(point1);
                    plpartTopDimY_.Add(point1);
                    plpartBotDimY_.Add(point1);

                    //plpartLeftDimX.Add(point2);
                    //plpartRightDimX.Add(point2);
                    //plpartLeftDimX_.Add(point2);
                    //plpartRightDimX_.Add(point2);

                    //plpartTopDimY.Add(point2);
                    //plpartBotDimY.Add(point2);
                    //plpartTopDimY_.Add(point2);
                    //plpartBotDimY_.Add(point2);
                    foreach (tsm.Part tamThayBoltTron in platesTamThayBoltTron) //Duyet tat cả các part cua  partsTamThayBoltTron
                    {
                        tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                        tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                        tsd.PointList pointsDimBoltInternal_X = new PointList();
                        tsd.PointList pointsDimBoltInternal_Y = new PointList();

                        List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                        Part_Edge partEdgeTamThayBoltTron = new Part_Edge(viewcur, tamThayBoltTron); // ds chưa tất cả điểm part edge cua  partEdgeTamThayBoltTron
                        List<t3d.Point> listpartEdgeTamThayBoltTron = partEdgeTamThayBoltTron.List_Edge;//ds điểm part edge cua tamThayBoltTron
                        t3d.Point xMinPart = partEdgeTamThayBoltTron.PointXmin0; // điểm X nhỏ nhất của part, để đo bolt vào part
                        t3d.Point xMaxPart = partEdgeTamThayBoltTron.PointXmax0; // điểm X lớn nhất của part, để đo bolt vào part
                        t3d.Point yMinPart = partEdgeTamThayBoltTron.PointYmin0; // điểm Y nhỏ nhất của part, để đo bolt vào part
                        t3d.Point yMaxPart = partEdgeTamThayBoltTron.PointYmax0; // điểm Y lớn nhất của part, để đo bolt vào part

                        foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts()) //Phần này cần xử lý lại để lấy được điểm bolt mong muốn
                        {
                            Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                            try
                            {
                                t3d.Point minY = tinh_Toan_Bolt.PointListBolt_XY_minmaxY[0]; //lấy điểm bolt có Y nhỏ nhất
                                t3d.Point maxY = tinh_Toan_Bolt.PointListBolt_XY_minmaxY[1]; //lấy điểm bolt có Y lớn nhất
                                t3d.Point minX = tinh_Toan_Bolt.PointListBolt_XY_minmaxX[0]; //lấy điểm bolt có X nhỏ nhất
                                t3d.Point maxX = tinh_Toan_Bolt.PointListBolt_XY_minmaxX[1]; //lấy điểm bolt có X lớn nhất
                                plpartLeftDimX.Add(minY);
                                plpartRightDimX.Add(minY);
                                plpartLeftDimX_.Add(maxY);
                                plpartRightDimX_.Add(maxY);
                                plpartTopDimY.Add(minX);
                                plpartBotDimY.Add(minX);
                                plpartTopDimY_.Add(maxX);
                                plpartBotDimY_.Add(maxX);
                            }
                            catch
                            {
                                plpartLeftDimX.Add(xMinPart);
                                plpartRightDimX.Add(xMaxPart);
                                plpartLeftDimX_.Add(xMinPart);
                                plpartRightDimX_.Add(xMaxPart);
                                plpartTopDimY.Add(yMaxPart);
                                plpartBotDimY.Add(yMinPart);
                                plpartTopDimY_.Add(yMaxPart);
                                plpartBotDimY_.Add(yMinPart);
                            }
                            //****phần dưới này cần thay bằng 1 clash: đầu vào là tsm.BoltGroup, đầu ra là dim các điểm của bolt đó vào part****
                            foreach (t3d.Point p in bgr.BoltPositions)
                            {
                                pointsDimBoltInternal_X.Add(p);
                                pointsDimBoltInternal_Y.Add(p);
                            }
                            try
                            {
                                if (chbox_DimInternalPartInAssembly.Checked)
                                {
                                    if (point1.Y < point3.Y)
                                        pointsDimBoltInternal_Y.Add(yMinPart);
                                    else
                                        pointsDimBoltInternal_Y.Add(yMaxPart);
                                    createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimBoltInternal_X, cb_DimentionAttributes.Text, 200); //tạo dim nội bộ bolts
                                    createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimBoltInternal_Y, cb_DimentionAttributes.Text, 200); //tạo dim nội bộ bolts
                                    pointsDimBoltInternal_X.Clear();
                                    pointsDimBoltInternal_Y.Clear();
                                }
                            }
                            catch
                            { }
                        }


                    } //Duyet tat cả các part cua  partsTamThayBoltTron
                    foreach (tsm.Part tamDung in platesTamDung) //Duyet tat cả các part cua  partsTamDung
                    {
                        Part_Edge part_EdgeTamDung = new Part_Edge(viewcur, tamDung); // ds chưa tất cả điểm part edge cua  tamDung
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                        t3d.Point pointTamDungXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                        t3d.Point pointTamDungXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                        t3d.Point pointTamDungXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                        t3d.Point pointTamDungXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải

                        plpartLeftDimX.Add(pointTamDungXminYmax);
                        plpartRightDimX.Add(pointTamDungXmaxYmax);
                        plpartLeftDimX_.Add(pointTamDungXminYmin);
                        plpartRightDimX_.Add(pointTamDungXmaxYmin);

                    } //Duyet tat cả các part cua  partsTamDung
                    foreach (tsm.Part tamNgang in platesTamNgang) //Duyet tat cả các part cua  partsTamNgang
                    {
                        Part_Edge part_EdgeTamNgang = new Part_Edge(viewcur, tamNgang); // ds chưa tất cả điểm part edge cua  partsTamNgang
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamNgang.List_Edge;//ds điểm của part 
                        t3d.Point pointTamNgangXminYmin = part_EdgeTamNgang.PointXminYmin; // góc dưới bên traí
                        t3d.Point pointTamNgangXminYmax = part_EdgeTamNgang.PointXminYmax; // góc trên bên trái
                        t3d.Point pointTamNgangXmaxYmax = part_EdgeTamNgang.PointXmaxYmax; // góc trên bên phải
                        t3d.Point pointTamNgangXmaxYmin = part_EdgeTamNgang.PointXmaxYmin; // góc dưới bên phải

                        plpartTopDimY.Add(pointTamNgangXmaxYmax);
                        plpartBotDimY.Add(pointTamNgangXmaxYmin);
                        plpartTopDimY_.Add(pointTamNgangXminYmax);
                        plpartBotDimY_.Add(pointTamNgangXminYmin);
                    } //Duyet tat cả các part cua  partsTamNgang
                    foreach (var tamNghiengPhai in platesNghiengPhai)//Duyet tat cả các part cua  partsNghiengPhai
                    {
                        Part_Edge part_EdgeNghiengPhai = new Part_Edge(viewcur, tamNghiengPhai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengPhai.List_Edge;//ds điểm của part 
                        t3d.Point tamNghiengPhaiXmin = part_EdgeNghiengPhai.PointXmin0;
                        t3d.Point tamNghiengPhaiXmax = part_EdgeNghiengPhai.PointXmax0;
                        t3d.Point tamNghiengPhaiYmin = part_EdgeNghiengPhai.PointYmin0;
                        t3d.Point tamNghiengPhaiYmax = part_EdgeNghiengPhai.PointYmax0;

                        plpartLeftDimX.Add(tamNghiengPhaiYmax);
                        plpartRightDimX.Add(tamNghiengPhaiYmax);
                        plpartLeftDimX_.Add(tamNghiengPhaiYmin);
                        plpartRightDimX_.Add(tamNghiengPhaiYmin);

                        plpartTopDimY.Add(tamNghiengPhaiXmax);
                        plpartBotDimY.Add(tamNghiengPhaiXmax);
                        plpartTopDimY_.Add(tamNghiengPhaiXmin);
                        plpartBotDimY_.Add(tamNghiengPhaiXmin);
                        if (point3.Y > point1.Y)
                        {
                            tsd.PointList listDoNghieng = new PointList(); //Tập hợp điểm DIM mép trái PL phuong X
                            listDoNghieng.Add(tamNghiengPhaiYmax);
                            listDoNghieng.Add(tamNghiengPhaiXmin);
                            int spaceDimX1 = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                            createDimension.CreateStraightDimensionSet_X(viewbase, listDoNghieng, cb_DimentionAttributes.Text, spaceDimX1);
                        }
                        else if (point3.Y < point1.Y)
                        {
                            tsd.PointList listDoNghieng = new PointList(); //Tập hợp điểm DIM mép trái PL phuong X
                            listDoNghieng.Add(tamNghiengPhaiYmin);
                            listDoNghieng.Add(tamNghiengPhaiXmax);
                            int spaceDimX1 = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                            createDimension.CreateStraightDimensionSet_X_(viewbase, listDoNghieng, cb_DimentionAttributes.Text, spaceDimX1);
                        }

                    }//Duyet tat cả các part cua  partsNghiengPhai
                    foreach (var tamNghiengTrai in platesNghiengTrai)//Duyet tat cả các part cua  partsNghiengTrai
                    {
                        Part_Edge part_EdgeNghiengTrai = new Part_Edge(viewcur, tamNghiengTrai); // ds chưa tất cả điểm part edge cua  partsNghiengTrai
                        List<t3d.Point> listpartEdgeNghiengTrai = part_EdgeNghiengTrai.List_Edge;//ds điểm của part 
                        t3d.Point tamNghiengTraiXmin = part_EdgeNghiengTrai.PointXmin0;
                        t3d.Point tamNghiengTraiXmax = part_EdgeNghiengTrai.PointXmax0;
                        t3d.Point tamNghiengTraiYmin = part_EdgeNghiengTrai.PointYmin0;
                        t3d.Point tamNghiengTraiYmax = part_EdgeNghiengTrai.PointYmax0;

                        plpartLeftDimX.Add(tamNghiengTraiYmax);
                        plpartRightDimX.Add(tamNghiengTraiYmax);
                        plpartLeftDimX_.Add(tamNghiengTraiYmin);
                        plpartRightDimX_.Add(tamNghiengTraiYmin);

                        plpartTopDimY.Add(tamNghiengTraiXmax);
                        plpartBotDimY.Add(tamNghiengTraiXmax);
                        plpartTopDimY_.Add(tamNghiengTraiXmin);
                        plpartBotDimY_.Add(tamNghiengTraiXmin);
                        if (point3.Y > point1.Y)
                        {
                            tsd.PointList listDoNghieng = new PointList(); //Tập hợp điểm DIM mép trái PL phuong X
                            listDoNghieng.Add(tamNghiengTraiYmax);
                            listDoNghieng.Add(tamNghiengTraiXmax);
                            int spaceDimX1 = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                            createDimension.CreateStraightDimensionSet_X(viewbase, listDoNghieng, cb_DimentionAttributes.Text, spaceDimX1);
                        }
                        else if (point3.Y < point1.Y)
                        {
                            tsd.PointList listDoNghieng = new PointList(); //Tập hợp điểm DIM mép trái PL phuong X
                            listDoNghieng.Add(tamNghiengTraiYmin);
                            listDoNghieng.Add(tamNghiengTraiXmin);
                            int spaceDimX1 = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                            createDimension.CreateStraightDimensionSet_X_(viewbase, listDoNghieng, cb_DimentionAttributes.Text, spaceDimX1);
                        }
                    }//Duyet tat cả các part cua  partsNghiengTrai
                    foreach (var partIDung in partsIDung) //Part U, I  ĐỨNG
                    {
                        string partProfileType = null;
                        partIDung.GetReportProperty("PROFILE_TYPE", ref partProfileType);
                        if (partProfileType == "U" || partProfileType == "L") // nếu là thép U thì lấy referenline
                        {
                            ArrayList arrayListRefLine = partIDung.GetReferenceLine(true);
                            List<t3d.Point> listRefLine = new List<t3d.Point> { arrayListRefLine[0] as t3d.Point, arrayListRefLine[1] as t3d.Point };
                            t3d.Point pointRefLineXmin = listRefLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointRefLineXmax = listRefLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointRefLineYmin = listRefLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointRefLineYmax = listRefLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointRefLineYmax);
                            plpartRightDimX.Add(pointRefLineYmax);
                            plpartLeftDimX_.Add(pointRefLineYmin);
                            plpartRightDimX_.Add(pointRefLineYmin);
                        }
                        else
                        {
                            ArrayList arrayListCenterLine = partIDung.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointCenterLineYmax);
                            plpartRightDimX.Add(pointCenterLineYmax);
                            plpartLeftDimX_.Add(pointCenterLineYmin);
                            plpartRightDimX_.Add(pointCenterLineYmin);
                        }
                    } //Part U, I  ĐỨNG
                    foreach (var partINgang in partsINgang) //Part I  NGANG
                    {
                        string partProfileType = null;
                        partINgang.GetReportProperty("PROFILE_TYPE", ref partProfileType);
                        if (partProfileType == "U" || partProfileType == "L") // nếu là thép U thì lấy referenline
                        {
                            ArrayList arrayListRefLine = partINgang.GetReferenceLine(true);
                            List<t3d.Point> listRefLine = new List<t3d.Point> { arrayListRefLine[0] as t3d.Point, arrayListRefLine[1] as t3d.Point };
                            t3d.Point pointRefLineXmin = listRefLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointRefLineXmax = listRefLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointRefLineYmin = listRefLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointRefLineYmax = listRefLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointRefLineYmax);
                            plpartRightDimX.Add(pointRefLineYmax);
                            plpartLeftDimX_.Add(pointRefLineYmin);
                            plpartRightDimX_.Add(pointRefLineYmin);
                        }
                        else
                        {
                            ArrayList arrayListCenterLine = partINgang.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointCenterLineYmax);
                            plpartRightDimX.Add(pointCenterLineYmax);
                            plpartLeftDimX_.Add(pointCenterLineYmin);
                            plpartRightDimX_.Add(pointCenterLineYmin);
                        }
                    }  //Part U, I  NGANG
                    foreach (var partINghiengTrai in partsINghiengTrai) //Part I  NGHIÊNG TRÁI
                    {
                        //MessageBox.Show("TRA");
                        ArrayList arrayListCenterLine = partINghiengTrai.GetCenterLine(true);
                        List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                        t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                        t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                        t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                        t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                        plpartLeftDimX.Add(pointCenterLineYmin);
                        plpartRightDimX.Add(pointCenterLineYmin);
                        plpartLeftDimX_.Add(pointCenterLineYmax);
                        plpartRightDimX_.Add(pointCenterLineYmax);

                        plpartTopDimY.Add(pointCenterLineXmin);
                        plpartBotDimY.Add(pointCenterLineXmin);
                        plpartTopDimY_.Add(pointCenterLineXmax);
                        plpartBotDimY_.Add(pointCenterLineXmax);

                        if (chbox_DimInternalPartInAssembly.Checked)
                        {
                            tsd.PointList plCenterPartINghiengTrai = new PointList(); //Tập hợp 2 điểm center line
                            plCenterPartINghiengTrai.Add(pointCenterLineYmin);
                            plCenterPartINghiengTrai.Add(pointCenterLineYmax);
                            createDimension.CreateStraightDimensionSet_X(viewbase, plCenterPartINghiengTrai, cb_DimentionAttributes.Text, 200); //tạo dim nội bộ bolts
                        }

                    }
                    foreach (var partINghiengPhai in partsINghiengPhai) //Part I  NGHIÊNG PHẢI
                    {
                        //MessageBox.Show("PHA");
                        ArrayList arrayListCenterLine = partINghiengPhai.GetCenterLine(true);
                        List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                        t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                        t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                        t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                        t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                        plpartLeftDimX.Add(pointCenterLineYmin);
                        plpartRightDimX.Add(pointCenterLineYmin);
                        plpartLeftDimX_.Add(pointCenterLineYmax);
                        plpartRightDimX_.Add(pointCenterLineYmax);

                        plpartTopDimY.Add(pointCenterLineXmin);
                        plpartBotDimY.Add(pointCenterLineXmin);
                        plpartTopDimY_.Add(pointCenterLineXmax);
                        plpartBotDimY_.Add(pointCenterLineXmax);

                        if (chbox_DimInternalPartInAssembly.Checked)
                        {
                            tsd.PointList plCenterPartINghiengTrai = new PointList(); //Tập hợp 2 điểm center line
                            plCenterPartINghiengTrai.Add(pointCenterLineYmin);
                            plCenterPartINghiengTrai.Add(pointCenterLineYmax);
                            createDimension.CreateStraightDimensionSet_X(viewbase, plCenterPartINghiengTrai, cb_DimentionAttributes.Text, 200); //tạo dim nội bộ bolts
                        }
                    }
                    foreach (var partLUsection in partsLUsection) //Part L, U section
                    {
                        //MessageBox.Show("PHA");
                        ArrayList arrayListRefLine = partLUsection.GetReferenceLine(true);
                        List<t3d.Point> listRefLine = new List<t3d.Point> { arrayListRefLine[0] as t3d.Point, arrayListRefLine[1] as t3d.Point };
                        t3d.Point pointRefLineX = listRefLine.OrderBy(point => point.X).ToList()[0];

                        plpartLeftDimX.Add(pointRefLineX);
                        plpartRightDimX.Add(pointRefLineX);
                        plpartLeftDimX_.Add(pointRefLineX);
                        plpartRightDimX_.Add(pointRefLineX);

                        plpartTopDimY.Add(pointRefLineX);
                        plpartBotDimY.Add(pointRefLineX);
                        plpartTopDimY_.Add(pointRefLineX);
                        plpartBotDimY_.Add(pointRefLineX);

                        if (chbox_DimInternalPartInAssembly.Checked)
                        {
                            //tạm thời chưa làm gì
                        }
                    }

                    int spaceDimX = Convert.ToInt32(Math.Abs(plpartLeftDimX[0].Y - point3.Y));
                    int spaceDimY = Convert.ToInt32(Math.Abs(plpartTopDimY[0].X - point3.X));
                    try
                    {
                        if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1)// Nếu chọn dim bên trái tấm
                        {
                            if (point3.Y > point1.Y)
                                createDimension.CreateStraightDimensionSet_X(viewbase, plpartLeftDimX, cb_DimAttDimPartBoltAssembly.Text, spaceDimX);
                            else if (point3.Y < point1.Y)
                                createDimension.CreateStraightDimensionSet_X_(viewbase, plpartLeftDimX_, cb_DimAttDimPartBoltAssembly.Text, spaceDimX);
                            Operation.DisplayPrompt("Sorry!");
                        }
                        else if (imcb_DimensionTo.SelectedIndex == 1) // Nếu chọn dim bên Phải tấm 
                        {
                            if (point3.Y > point1.Y)
                                createDimension.CreateStraightDimensionSet_X(viewbase, plpartRightDimX, cb_DimAttDimPartBoltAssembly.Text, spaceDimX);
                            else if (point3.Y < point1.Y)
                                createDimension.CreateStraightDimensionSet_X_(viewbase, plpartRightDimX_, cb_DimAttDimPartBoltAssembly.Text, spaceDimX);
                        }
                        else if (imcb_DimensionTo.SelectedIndex == 2) // Nếu chọn dim 2 bên tấm 
                            Operation.DisplayPrompt("Sorry!");
                    }
                    catch
                    { }
                }

                else if (partDistribution.havePart == false && partDistribution.haveBolt == true) //Phần này để dim cho BOLTs khi người dùng chỉ chọn BOLT trên bản vẽ.
                {
                    //MessageBox.Show("dim bolt ne");
                    List<List<t3d.Point>> listListPointsBolt_XY = partDistribution.listListPointsBolt_XY;
                    List<List<t3d.Point>> listListPointsBolt_X = partDistribution.listListPointsBolt_X;
                    List<List<t3d.Point>> listListPointsBolt_Y = partDistribution.listListPointsBolt_Y;
                    PointList pointsDimBolt_Min = new PointList();
                    PointList pointsDimBolt_Max = new PointList();

                    pointsDimBolt_Min.Add(point1);
                    pointsDimBolt_Max.Add(point1);

                    foreach (List<t3d.Point> listPointsBolt_XY in listListPointsBolt_XY)
                    {
                        List<t3d.Point> points = new List<t3d.Point>();
                        List<t3d.Point> pointsDes = new List<t3d.Point>();
                        PointList pointsDimBoltInternal = new PointList();
                        points = listPointsBolt_XY.OrderBy(p => p.X).ToList();
                        pointsDimBolt_Min.Add(points[0]);
                        pointsDes = listPointsBolt_XY.OrderByDescending(p => p.X).ToList();
                        pointsDimBolt_Max.Add(pointsDes[0]);
                        foreach (t3d.Point point in listPointsBolt_XY)
                        {

                            pointsDimBoltInternal.Add(point);
                        }
                        if (pointsDimBoltInternal.Count != 0)
                            try
                            {
                                createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimBoltInternal, cb_DimentionAttributes.Text, 100); //tạo dim nội bộ bolts
                                createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimBoltInternal, cb_DimentionAttributes.Text, 100); //tạo dim nội bộ bolts
                            }
                            catch { }
                    } //Tính toán Lấy điểm bolt của bolt XY
                    foreach (List<t3d.Point> listPointsBolt_Y in listListPointsBolt_Y)
                    {
                        List<t3d.Point> points = new List<t3d.Point>();
                        List<t3d.Point> pointsDes = new List<t3d.Point>();
                        PointList pointsDimBoltInternal = new PointList(); //dim tung bolt trong boltgroup
                        points = listPointsBolt_Y.OrderBy(p => p.X).ToList();
                        pointsDimBolt_Min.Add(points[0]);
                        pointsDes = listPointsBolt_Y.OrderByDescending(p => p.X).ToList();
                        pointsDimBolt_Max.Add(pointsDes[0]);
                        foreach (t3d.Point point in listPointsBolt_Y)
                        {
                            pointsDimBoltInternal.Add(point);
                        }
                        if (pointsDimBoltInternal.Count != 0)
                            try
                            {
                                createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimBoltInternal, cb_DimentionAttributes.Text, 100); //tạo dim nội bộ bolts, do bolt Y nen chi co dim X
                            }
                            catch { }
                    } //Tính toán  lấy điểm bolt của bolt XY

                    int spaceDim = Convert.ToInt32(Math.Abs(point3.Y - point1.Y));
                    if (point1.X < pointsDimBolt_Min[1].X) //Nếu mà pick1 nàm bên trái bolt nhỏ nhất, thì dim vào bolt trái của nhóm bolt
                    {
                        if (point3.Y > point1.Y)
                            createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimBolt_Min, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                        else
                            createDimension.CreateStraightDimensionSet_X_(viewbase, pointsDimBolt_Min, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                    }
                    else
                    {
                        if (point3.Y > point1.Y)
                            createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimBolt_Max, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                        else
                            createDimension.CreateStraightDimensionSet_X_(viewbase, pointsDimBolt_Max, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                    }
                }
            } //NEU CHON DIM PARTS OR BOLTS ASSEMBLY
            if (chbox_AddPartMark.Checked)
                Add_Mark();
            drawingHandler.GetActiveDrawing().CommitChanges();
        } // DIM GA THEO PHUONG X

        private void btn_DimGA_Y_Click(object sender, EventArgs e) // DIM GA THEO PHUONG Y
        {
            try
            {
                // Lấy đối tượng được chọn trên bản vẽ
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                CreateDimension createDimension = new CreateDimension(); // tạo 1 createDimension xem thêm class CreateDimension
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                tsm.Part mainPart = null;
                PartDistribution partDistribution = null;
                if (actived_dr is tsd.AssemblyDrawing)
                {
                    tsd.AssemblyDrawing assemblyDrawing = actived_dr as tsd.AssemblyDrawing; //Lấy bản vẽ assembly
                    tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                    mainPart = assembly.GetMainPart() as tsm.Part;
                    partDistribution = new PartDistribution(mainPart, drobjenum);
                }
                tsdui.Picker c_picker = drawingHandler.GetPicker();
                var pick = c_picker.PickPoint("Pick first dimension point");
                //Lấy tọa độ điểm pick (điểm gốc)
                t3d.Point point1 = pick.Item1;

                var pick3 = c_picker.PickPoint("Pick location for dimension");
                t3d.Point point3 = pick3.Item1;
                //Lấy view hiện hành thông qua điểm vừa pick
                tsd.View viewcur = pick.Item2 as tsd.View;
                tsd.ViewBase viewbase = pick.Item2;
                tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                wph.SetCurrentTransformationPlane(viewplane);

                tsd.PointList pointListDimRef = new tsd.PointList() { point1 };
                tsd.PointList pointListDimCenter = new tsd.PointList() { point1 };

                tsd.PointList pointListDimEde_X = new PointList() { point1 };//Tập hợp điểm X bên trên
                tsd.PointList pointListDimEde_X_ = new PointList() { point1 };//Tập hợp điểm X bên dưới
                tsd.PointList pointListDimEde_Y = new PointList() { point1 };//Tập hợp điểm Y bên phải
                tsd.PointList pointListDimEde_Y_ = new PointList() { point1 }; //Tập hợp điểm Y bên trái

                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drPart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                        Part_Edge part_edge = new Part_Edge(viewcur, mPart);

                        t3d.Point maxX = part_edge.PointXmax0;
                        t3d.Point minX = part_edge.PointXmin0;
                        t3d.Point maxY = part_edge.PointYmax0;
                        t3d.Point minY = part_edge.PointYmin0;
                        pointListDimEde_X.Add(part_edge.PointXminYmax);
                        pointListDimEde_X.Add(part_edge.PointXmaxYmax);
                        pointListDimEde_X_.Add(part_edge.PointXminYmin);
                        pointListDimEde_X_.Add(part_edge.PointXmaxYmin);
                        pointListDimEde_Y.Add(part_edge.PointXmaxYmax);
                        pointListDimEde_Y.Add(part_edge.PointXmaxYmin);
                        pointListDimEde_Y_.Add(part_edge.PointXminYmax);
                        pointListDimEde_Y_.Add(part_edge.PointXminYmin);
                        ArrayList listRefLine = mPart.GetReferenceLine(false);
                        ArrayList listCenterLine = mPart.GetCenterLine(true);
                        foreach (t3d.Point p in listRefLine)
                        {
                            pointListDimRef.Add(p);
                        }

                        foreach (t3d.Point p in listCenterLine)
                        {
                            pointListDimCenter.Add(p);
                        }
                    }
                }
                drobjenum.Reset();

                if (imcb_SelectDimRefCenEdge.SelectedIndex == 1) //NEU CHON DIM REFERENCE LINE
                {
                    int spaceDim = Convert.ToInt32(Math.Abs(pointListDimRef[0].X - point3.X)); //Tinh khoang cach tu diem dim dau tien, đến điểm đặt dim
                    if (point1.X < point3.X)
                    {
                        createDimension.CreateStraightDimensionSet_Y(viewbase, pointListDimRef, cb_DimentionAttributes.Text, spaceDim);
                    }
                    else
                    {
                        createDimension.CreateStraightDimensionSet_Y_(viewbase, pointListDimRef, cb_DimentionAttributes.Text, spaceDim);
                    }
                }
                else if (imcb_SelectDimRefCenEdge.SelectedIndex == 2) //NEU CHON DIM CENTER LINE
                {
                    int spaceDim = Convert.ToInt32(Math.Abs(pointListDimCenter[0].X - point3.X)); //Tinh khoang cach tu diem dim dau tien, đến điểm đặt dim
                    if (point1.X < point3.X)
                    {
                        createDimension.CreateStraightDimensionSet_Y(viewbase, pointListDimCenter, cb_DimentionAttributes.Text, spaceDim);
                    }
                    else
                    {
                        createDimension.CreateStraightDimensionSet_Y_(viewbase, pointListDimCenter, cb_DimentionAttributes.Text, spaceDim);
                    }
                }
                else if (imcb_SelectDimRefCenEdge.SelectedIndex == 3) //NEU CHON DIM EDGE
                {
                    int spaceDim = Convert.ToInt32(Math.Abs(pointListDimEde_Y[0].X - point3.X)); //Tinh khoang cach tu diem dim dau tien, đến điểm đặt dim
                    if (point1.X < point3.X)
                    {
                        createDimension.CreateStraightDimensionSet_Y(viewbase, pointListDimEde_Y, cb_DimentionAttributes.Text, spaceDim);
                    }
                    else
                    {
                        createDimension.CreateStraightDimensionSet_Y_(viewbase, pointListDimEde_Y_, cb_DimentionAttributes.Text, spaceDim);
                    }
                }
                else if (imcb_SelectDimRefCenEdge.SelectedIndex == 0 || imcb_SelectDimRefCenEdge.SelectedIndex == -1) //NEU CHON DIM PARTS OR BOLTS ASSEMBLY
                {
                    if (partDistribution.havePart == true)
                    {
                        List<tsm.Part> partsTamThayBoltTron = partDistribution.ListTamThayBoltTron;
                        List<tsm.Part> partsTamDung = partDistribution.ListTamDung;
                        List<tsm.Part> partsTamNgang = partDistribution.ListTamNgang;
                        List<tsm.Part> partsNghiengPhai = partDistribution.ListTamNghiengPhai;
                        List<tsm.Part> partsNghiengTrai = partDistribution.ListTamNghiengTrai;
                        List<tsm.Part> partsIDung = partDistribution.ListPartIDung;
                        List<tsm.Part> partsINgang = partDistribution.ListPartINgang;
                        List<tsm.Part> partsINghiengTrai = partDistribution.ListPartINghiengTrai;
                        List<tsm.Part> partsINghiengPhai = partDistribution.ListPartINghiengPhai;

                        tsd.PointList plpartLeftDimX = new PointList(); //Tập hợp điểm DIM mép trái PL phuong X
                        tsd.PointList plpartRightDimX = new PointList(); //Tập hợp điểm DIM mép phai PL phuong X
                        tsd.PointList plpartLeftDimX_ = new PointList();//Tập hợp điểm DIM mép trái PL phuong X_
                        tsd.PointList plpartRightDimX_ = new PointList();//Tập hợp điểm DIM mép phai PL phuong X_

                        tsd.PointList plpartTopDimY = new PointList(); //Tập hợp điểm DIM mép tren PL phuong Y
                        tsd.PointList plpartBotDimY = new PointList(); //Tập hợp điểm DIM mép duoi PL phuong Y
                        tsd.PointList plpartTopDimY_ = new PointList();//Tập hợp điểm DIM mép tren PL phuong Y_
                        tsd.PointList plpartBotDimY_ = new PointList();//Tập hợp điểm DIM mép duoi PL phuong Y_

                        plpartLeftDimX.Add(point1);
                        plpartRightDimX.Add(point1);
                        plpartLeftDimX_.Add(point1);
                        plpartRightDimX_.Add(point1);

                        plpartTopDimY.Add(point1);
                        plpartBotDimY.Add(point1);
                        plpartTopDimY_.Add(point1);
                        plpartBotDimY_.Add(point1);

                        foreach (tsm.Part tamThayBoltTron in partsTamThayBoltTron) //Duyet tat cả các part cua  partsTamThayBoltTron
                        {
                            tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                            tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                            List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                            Part_Edge partEdgeTamThayBoltTron = new Part_Edge(viewcur, tamThayBoltTron); // ds chưa tất cả điểm part edge cua  partEdgeTamThayBoltTron
                            List<t3d.Point> listpartEdgeTamThayBoltTron = partEdgeTamThayBoltTron.List_Edge;//ds điểm part edge cua tamThayBoltTron
                            foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts()) //Phần này cần xử lý lại để lấy được điểm bolt mong muốn
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    plpartLeftDimX.Add(p);
                                    plpartRightDimX.Add(p);
                                    plpartLeftDimX_.Add(p);
                                    plpartRightDimX_.Add(p);

                                    plpartTopDimY.Add(p);
                                    plpartBotDimY.Add(p);
                                    plpartTopDimY_.Add(p);
                                    plpartBotDimY_.Add(p);

                                    break;
                                }
                            }
                            if (chbox_DimInternalPartInAssembly.Checked)
                            {

                            }
                        } //Duyet tat cả các part cua  partsTamThayBoltTron
                        foreach (tsm.Part tamDung in partsTamDung) //Duyet tat cả các part cua  partsTamDung
                        {
                            Part_Edge part_EdgeTamDung = new Part_Edge(viewcur, tamDung); // ds chưa tất cả điểm part edge cua  tamDung
                            List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                            t3d.Point pointTamDungXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                            t3d.Point pointTamDungXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                            t3d.Point pointTamDungXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                            t3d.Point pointTamDungXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải

                            plpartLeftDimX.Add(pointTamDungXminYmax);
                            plpartRightDimX.Add(pointTamDungXmaxYmax);
                            plpartLeftDimX_.Add(pointTamDungXminYmin);
                            plpartRightDimX_.Add(pointTamDungXmaxYmin);

                        } //Duyet tat cả các part cua  partsTamDung
                        foreach (tsm.Part tamNgang in partsTamNgang) //Duyet tat cả các part cua  partsTamNgang
                        {
                            Part_Edge part_EdgeTamNgang = new Part_Edge(viewcur, tamNgang); // ds chưa tất cả điểm part edge cua  partsTamNgang
                            List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamNgang.List_Edge;//ds điểm của part 
                            t3d.Point pointTamNgangXminYmin = part_EdgeTamNgang.PointXminYmin; // góc dưới bên traí
                            t3d.Point pointTamNgangXminYmax = part_EdgeTamNgang.PointXminYmax; // góc trên bên trái
                            t3d.Point pointTamNgangXmaxYmax = part_EdgeTamNgang.PointXmaxYmax; // góc trên bên phải
                            t3d.Point pointTamNgangXmaxYmin = part_EdgeTamNgang.PointXmaxYmin; // góc dưới bên phải

                            plpartTopDimY.Add(pointTamNgangXmaxYmax);
                            plpartBotDimY.Add(pointTamNgangXmaxYmin);
                            plpartTopDimY_.Add(pointTamNgangXminYmax);
                            plpartBotDimY_.Add(pointTamNgangXminYmin);
                        } //Duyet tat cả các part cua  partsTamDung
                        foreach (var tamNghiengPhai in partsNghiengPhai)//Duyet tat cả các part cua  partsNghiengPhai
                        {
                            Part_Edge part_EdgeNghiengPhai = new Part_Edge(viewcur, tamNghiengPhai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                            List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengPhai.List_Edge;//ds điểm của part 
                            t3d.Point tamNghiengPhaiXmin = part_EdgeNghiengPhai.PointXmin0;
                            t3d.Point tamNghiengPhaiXmax = part_EdgeNghiengPhai.PointXmax0;
                            t3d.Point tamNghiengPhaiYmin = part_EdgeNghiengPhai.PointYmin0;
                            t3d.Point tamNghiengPhaiYmax = part_EdgeNghiengPhai.PointYmax0;

                            plpartLeftDimX.Add(tamNghiengPhaiYmin);
                            plpartRightDimX.Add(tamNghiengPhaiYmin);
                            plpartLeftDimX_.Add(tamNghiengPhaiYmax);
                            plpartRightDimX_.Add(tamNghiengPhaiYmax);

                            plpartTopDimY.Add(tamNghiengPhaiXmin);
                            plpartBotDimY.Add(tamNghiengPhaiXmin);
                            plpartTopDimY_.Add(tamNghiengPhaiXmax);
                            plpartBotDimY_.Add(tamNghiengPhaiXmax);
                        }//Duyet tat cả các part cua  partsNghiengPhai
                        foreach (var tamNghiengTrai in partsNghiengTrai)//Duyet tat cả các part cua  partsNghiengTrai
                        {
                            Part_Edge part_EdgeNghiengTrai = new Part_Edge(viewcur, tamNghiengTrai); // ds chưa tất cả điểm part edge cua  partsNghiengTrai
                            List<t3d.Point> listpartEdgeNghiengTrai = part_EdgeNghiengTrai.List_Edge;//ds điểm của part 
                            t3d.Point tamNghiengTraiXmin = part_EdgeNghiengTrai.PointXmin0;
                            t3d.Point tamNghiengTraiXmax = part_EdgeNghiengTrai.PointXmax0;
                            t3d.Point tamNghiengTraiYmin = part_EdgeNghiengTrai.PointYmin0;
                            t3d.Point tamNghiengTraiYmax = part_EdgeNghiengTrai.PointYmax0;

                            plpartLeftDimX.Add(tamNghiengTraiYmin);
                            plpartRightDimX.Add(tamNghiengTraiYmin);
                            plpartLeftDimX_.Add(tamNghiengTraiYmax);
                            plpartRightDimX_.Add(tamNghiengTraiYmax);

                            plpartTopDimY.Add(tamNghiengTraiXmin);
                            plpartBotDimY.Add(tamNghiengTraiXmin);
                            plpartTopDimY_.Add(tamNghiengTraiXmax);
                            plpartBotDimY_.Add(tamNghiengTraiXmax);
                        }//Duyet tat cả các part cua  partsNghiengTrai

                        foreach (var partIDung in partsIDung)//Part I  ĐỨNG
                        {
                            ArrayList arrayListCenterLine = partIDung.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointCenterLineYmax);
                            plpartRightDimX.Add(pointCenterLineYmax);
                            plpartLeftDimX_.Add(pointCenterLineYmin);
                            plpartRightDimX_.Add(pointCenterLineYmin);
                        }
                        foreach (var partINgang in partsINgang)//Part I  NGANG
                        {
                            ArrayList arrayListCenterLine = partINgang.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointCenterLineYmax);
                            plpartRightDimX.Add(pointCenterLineYmax);
                            plpartLeftDimX_.Add(pointCenterLineYmin);
                            plpartRightDimX_.Add(pointCenterLineYmin);

                            plpartTopDimY.Add(pointCenterLineXmax);
                            plpartBotDimY.Add(pointCenterLineXmax);
                            plpartTopDimY_.Add(pointCenterLineXmin);
                            plpartBotDimY_.Add(pointCenterLineXmin);
                        }
                        foreach (var partINghiengTrai in partsINghiengTrai) //Part I  NGHIÊNG TRÁI
                        {
                            //MessageBox.Show("TRA");
                            ArrayList arrayListCenterLine = partINghiengTrai.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointCenterLineYmin);
                            plpartRightDimX.Add(pointCenterLineYmin);
                            plpartLeftDimX_.Add(pointCenterLineYmax);
                            plpartRightDimX_.Add(pointCenterLineYmax);

                            plpartTopDimY.Add(pointCenterLineXmin);
                            plpartBotDimY.Add(pointCenterLineXmin);
                            plpartTopDimY_.Add(pointCenterLineXmax);
                            plpartBotDimY_.Add(pointCenterLineXmax);

                        }
                        foreach (var partINghiengPhai in partsINghiengPhai) //Part I  NGHIÊNG PHẢI
                        {
                            ArrayList arrayListCenterLine = partINghiengPhai.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            t3d.Point pointCenterLineYmin = listCenterLine.OrderBy(point => point.Y).ToList()[0];
                            t3d.Point pointCenterLineYmax = listCenterLine.OrderBy(point => point.Y).ToList()[1];
                            plpartLeftDimX.Add(pointCenterLineYmin);
                            plpartRightDimX.Add(pointCenterLineYmin);
                            plpartLeftDimX_.Add(pointCenterLineYmax);
                            plpartRightDimX_.Add(pointCenterLineYmax);

                            plpartTopDimY.Add(pointCenterLineXmin);
                            plpartBotDimY.Add(pointCenterLineXmin);
                            plpartTopDimY_.Add(pointCenterLineXmax);
                            plpartBotDimY_.Add(pointCenterLineXmax);
                        }
                        int spaceDimX = Convert.ToInt32(Math.Abs(plpartLeftDimX[0].Y - point3.Y));
                        int spaceDimY = Convert.ToInt32(Math.Abs(plpartTopDimY[0].X - point3.X));

                        if (point3.X > point1.X)
                            createDimension.CreateStraightDimensionSet_Y(viewbase, plpartTopDimY, cb_DimAttDimPartBoltAssembly.Text, spaceDimY);
                        else if (point3.X < point1.X)
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, plpartTopDimY_, cb_DimAttDimPartBoltAssembly.Text, spaceDimY);
                    }
                    else if (partDistribution.havePart == false && partDistribution.haveBolt == true) //Phần này để dim cho BOLTs khi người dùng chỉ chọn BOLT trên bản vẽ.
                    {
                        //MessageBox.Show("dim bolt ne");
                        List<List<t3d.Point>> listListPointsBolt_XY = partDistribution.listListPointsBolt_XY;
                        List<List<t3d.Point>> listListPointsBolt_X = partDistribution.listListPointsBolt_X;
                        List<List<t3d.Point>> listListPointsBolt_Y = partDistribution.listListPointsBolt_Y;
                        PointList pointsDimBolt_Min = new PointList();
                        PointList pointsDimBolt_Max = new PointList();

                        pointsDimBolt_Min.Add(point1);
                        pointsDimBolt_Max.Add(point1);

                        foreach (List<t3d.Point> listPointsBolt_XY in listListPointsBolt_XY)
                        {
                            List<t3d.Point> points = new List<t3d.Point>();    //list points bolt sắp xếp Y tăng dần
                            List<t3d.Point> pointsDes = new List<t3d.Point>(); //list points bolt sắp xếp Y giảm dần
                            PointList pointsDimBoltInternal = new PointList();
                            points = listPointsBolt_XY.OrderBy(p => p.Y).ToList();
                            pointsDimBolt_Min.Add(points[0]);
                            pointsDes = listPointsBolt_XY.OrderByDescending(p => p.Y).ToList();
                            pointsDimBolt_Max.Add(pointsDes[0]);
                            foreach (t3d.Point point in listPointsBolt_XY)
                            {
                                pointsDimBoltInternal.Add(point);
                            }
                            if (pointsDimBoltInternal.Count != 0)
                                try
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewbase, pointsDimBoltInternal, cb_DimentionAttributes.Text, 100); //tạo dim nội bộ bolts
                                    createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimBoltInternal, cb_DimentionAttributes.Text, 100); //tạo dim nội bộ bolts
                                }
                                catch { }
                        }
                        foreach (List<t3d.Point> listPointsBolt_X in listListPointsBolt_X)
                        {
                            List<t3d.Point> points = new List<t3d.Point>();    //list points bolt sắp xếp Y tăng dần
                            List<t3d.Point> pointsDes = new List<t3d.Point>(); //list points bolt sắp xếp Y giảm dần
                            PointList pointsDimBoltInternal = new PointList(); //dim tung bolt trong boltgroup
                            points = listPointsBolt_X.OrderBy(p => p.Y).ToList();
                            pointsDimBolt_Min.Add(points[0]);
                            pointsDes = listPointsBolt_X.OrderByDescending(p => p.Y).ToList();
                            pointsDimBolt_Max.Add(pointsDes[0]);
                            foreach (t3d.Point point in listPointsBolt_X)
                            {
                                pointsDimBoltInternal.Add(point);
                            }
                            if (pointsDimBoltInternal.Count != 0)
                                try
                                {
                                    createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimBoltInternal, cb_DimentionAttributes.Text, 100); //tạo dim nội bộ bolts, do bolt Y nen chi co dim X
                                }
                                catch { }
                        }
                        int spaceDim = Convert.ToInt32(Math.Abs(point3.X - point1.X));
                        if (point1.Y < pointsDimBolt_Min[1].Y) //Nếu mà pick1 nhỏ hơn nằm dưới điểm bolt nhỏ nhất, thì dim vào bolt dưới của nhóm bolt
                        {
                            if (point3.X < point1.X)
                                createDimension.CreateStraightDimensionSet_Y_(viewbase, pointsDimBolt_Min, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                            else
                                createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimBolt_Min, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                        }
                        else
                        {
                            if (point3.X >= point1.X)
                                createDimension.CreateStraightDimensionSet_Y_(viewbase, pointsDimBolt_Max, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                            else
                                createDimension.CreateStraightDimensionSet_Y(viewbase, pointsDimBolt_Max, cb_DimAttDimPartBoltAssembly.Text, spaceDim);
                        } //Nếu chọn dim theo Y
                    }
                } //NEU CHON DIM PARTS OR BOLTS ASSEMBLY
                drawingHandler.GetActiveDrawing().CommitChanges();
            }
            catch
            { }
        } // DIM GA THEO PHUONG Y

        private void btn_DimGA_Free_Click(object sender, EventArgs e) // DIM GA THEO PHUONG FREE
        {
            try
            {
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                //tsm.Part mainPart = null;
                //if (actived_dr is tsd.AssemblyDrawing)
                //{
                //    tsd.AssemblyDrawing assemblyDrawing = actived_dr as tsd.AssemblyDrawing; //Lấy bản vẽ assembly
                //    tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                //    mainPart = assembly.GetMainPart() as tsm.Part;
                //}
                // Lấy đối tượng được chọn trên bản vẽ
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                CreateDimension createDimension = new CreateDimension(); // tạo 1 createDimension xem thêm class CreateDimension
                tsdui.Picker c_picker = drawingHandler.GetPicker();
                var pick = c_picker.PickPoint("Chon diem thu nhat");
                //Lấy tọa độ điểm pick (điểm gốc)
                t3d.Point point1 = pick.Item1;
                //Lấy view hiện hành thông qua điểm vừa pick
                tsd.View viewcur = pick.Item2 as tsd.View;
                tsd.ViewBase viewbase = pick.Item2;
                tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                wph.SetCurrentTransformationPlane(viewplane);

                tsd.PointList pointListDimRef = new tsd.PointList() { point1 };
                tsd.PointList pointListDimCenter = new tsd.PointList() { point1 };

                tsd.PointList pointListDimEdge = new PointList() { point1 };//Tập hợp điểm 
                try
                {
                    var pick2 = c_picker.PickPoint("Chon diem thu hai");
                    t3d.Point point2 = pick2.Item1;
                    pointListDimRef.Add(point2);
                    pointListDimCenter.Add(point2);
                    pointListDimEdge.Add(point2);
                }
                catch
                { }
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drPart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mPart = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                        Part_Edge part_edge = new Part_Edge(viewcur, mPart);
                        List<t3d.Point> listPartEdge = part_edge.List_Edge;
                        t3d.Point maxX = part_edge.PointXmax0;
                        t3d.Point minX = part_edge.PointXmin0;
                        t3d.Point maxY = part_edge.PointYmax0;
                        t3d.Point minY = part_edge.PointYmin0;
                        pointListDimEdge.Add(part_edge.PointXminYmax);
                        pointListDimEdge.Add(part_edge.PointXmaxYmax);
                        pointListDimEdge.Add(part_edge.PointXminYmin);
                        pointListDimEdge.Add(part_edge.PointXmaxYmin);

                        ArrayList listRefLine = mPart.GetReferenceLine(false);
                        ArrayList listCenterLine = mPart.GetCenterLine(false);
                        foreach (t3d.Point p in listRefLine)
                        {
                            pointListDimRef.Add(p);
                        }

                        foreach (t3d.Point p in listCenterLine)
                        {
                            pointListDimCenter.Add(p);
                        }
                        //wph.SetCurrentTransformationPlane(new TransformationPlane());
                    }
                }
                //var pick3 = c_picker.PickPoint("Chon diem dat dim");
                //t3d.Point point3 = pick3.Item1;
                var pick4 = c_picker.PickPoint("Chon diem 1 xac dinh phuong");
                t3d.Point point4 = pick4.Item1;
                var pick5 = c_picker.PickPoint("Chon diem 2 xac dinh phuong");
                t3d.Point point5 = pick5.Item1;
                if (imcb_SelectDimRefCenEdge.SelectedIndex == 0 || imcb_SelectDimRefCenEdge.SelectedIndex == -1) //NEU CHON DIM REFERENCE LINE
                {
                    createDimension.CreateStraightDimensionSet_FX(point4, point5, viewbase, pointListDimRef, cb_DimentionAttributes.Text, 250);
                }
                else if (imcb_SelectDimRefCenEdge.SelectedIndex == 1) //NEU CHON DIM CENTER LINE
                {
                    createDimension.CreateStraightDimensionSet_FX(point4, point5, viewbase, pointListDimCenter, cb_DimentionAttributes.Text, 250);
                }
                else if (imcb_SelectDimRefCenEdge.SelectedIndex == 2) //NEU CHON DIM EDGE
                {
                    createDimension.CreateStraightDimensionSet_FX(point4, point5, viewbase, pointListDimEdge, cb_DimentionAttributes.Text, 250);
                }
                drawingHandler.GetActiveDrawing().CommitChanges();
            }
            catch
            { }
        } // DIM GA THEO PHUONG FREE

        public void Add_Mark()
        {
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
            //Create a script as string
            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                            "{" +
                            " public class Script" +
                                "{" +
                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                    " {" +
            @"akit.Callback(""acmd_create_marks_selected"", """", ""main_frame"");" +
            @"akit.PushButton(""smark_cancel"", ""smark_dial"");" +
             @"akit.PushButton(""pmark_cancel"",  ""pmark_dial"");" +
            "}}}";
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }

        // TAB ***********[DIMENSION TOOL]***********

        // TAB ***********[NUMBERING DRAWING]***********
        private void btn_change_Click(object sender, EventArgs e) //button Changing thông số bản vẽ (tab numbering)
        {
            try
            {
                tsd.Drawing MyDrawing = drawingHandler.GetActiveDrawing();
                string Dr_Prefix = txt_DrPrefix.Text;
                string Dr_Posfix = txt_DrPostfix.Text;
                tsd.DrawingEnumerator MyDrEnum = drawingHandler.GetDrawingSelector().GetSelected();
                int Xi = MyDrEnum.GetSize();
                //MessageBox.Show(Xi.ToString());
                tsd.Drawing Dr = null;
                DialogResult dialogResult = MessageBox.Show(string.Format("Are you sure do you want to change attributes in {0} drawing?", Xi), "Are you sure?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1)
                    #region WRITE TO: Approved By
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 1)
                    #region Tittle 1
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            Dr.Title1 = Dr_Prefix + Dr_Posfix;
                            Dr.Modify();
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 2)
                    #region Tittle 2
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            Dr.Title2 = Dr_Prefix + Dr_Posfix;
                            Dr.Modify();
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 3)
                    #region Tittle 3
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            Dr.Title3 = Dr_Prefix + Dr_Posfix;
                            Dr.Modify();
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 4)
                    #region WRITE TO: User Field 1
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 5)
                    #region WRITE TO: User Field 2
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 6)
                    #region WRITE TO: User Field 3
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 7)
                    #region WRITE TO: Name
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            Dr.Name = Dr_Prefix + Dr_Posfix;
                            Dr.Modify();
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 8)
                    #region WRITE TO: Checked By
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 9)
                    #region WRITE TO: ASSIGNED TO
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion

                    else if (comboBox1.SelectedIndex == 10)
                    #region WRITE TO: Comment
                    {
                        while (MyDrEnum.MoveNext())
                        {
                            Dr = MyDrEnum.Current as Drawing;
                            string DrMk = Dr.Mark;
                            Dr.SetUserProperty("comment", Dr_Prefix + Dr_Posfix);
                        }
                        if (Dr is GADrawing)
                        {
                            Operation.RunMacro("UDA_G_Modify.cs");
                        }
                        else
                        {
                            Operation.RunMacro("UDA_A&M_Modify.cs");
                        }
                        MessageBox.Show("Changing complete -- Done");
                    }
                    #endregion
                }
                else
                { }
            }
            catch
            {
                MessageBox.Show("Wrong Tekla Structures Version or Start Number empty.");
            }
        }

        private void btn_NumberingDrawing_Click(object sender, EventArgs e) //button NumberingDrawing thông số bản vẽ (tab numbering)
        {
            try
            {
                tsd.Drawing MyDrawing = drawingHandler.GetActiveDrawing();
                string Dr_Prefix = txt_DrPrefix.Text; // lấy dữ liệu từ ô Prefix.
                string Dr_Posfix = txt_DrPostfix.Text; // lấy dữ liệu từ ô Postfix.
                int Dr_Sn = Convert.ToInt32(txt_DrStartNumber.Text); // lấy dữ liệu từ ô start no.
                int numberingDigit = Convert.ToInt32(txt_Digit.Text);
                tsd.DrawingEnumerator MyDrEnum = drawingHandler.GetDrawingSelector().GetSelected();
                int Xi = MyDrEnum.GetSize();
                //MessageBox.Show(Xi.ToString());
                char[] alphabet = Enumerable.Range('A', 26).Select(x => (char)x).ToArray();
                tsd.Drawing Dr = null;
                DialogResult dialogResult = MessageBox.Show(string.Format("Are you sure do you want to change attributes in {0} drawing?", Xi), "Are you sure?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (chb_AtadNumberingStyle.Checked)
                    {
                        if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1) //nếu Write to không chọn, hoặc chọn Approved by
                        #region WRITE TO: Approved By
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 1)
                        #region Tittle 1
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.Title1 = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;
                                }

                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.Title1 = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                            Dr.Modify();
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 2)
                        #region Tittle 2
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.Title2 = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;
                                }

                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.Title2 = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                            Dr.Modify();
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 3)
                        #region Tittle 3
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.Title3 = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;
                                }

                                else if (DrMk.Contains(" - "))
                                {

                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.Title3 = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                            Dr.Modify();
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 4)
                        #region WRITE TO: User Field 1
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 5)
                        #region WRITE TO: User Field 2
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 6)
                        #region WRITE TO: User Field 3
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 7)
                        #region WRITE TO:Name
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.Name = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;
                                }

                                else if (DrMk.Contains(" - "))
                                {

                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.Name = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                            Dr.Modify();
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 8)
                        #region WRITE TO: Checked By
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }

                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 9)
                        #region WRITE TO: Assigned To
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }

                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 10)
                        #region WRITE TO: Comment
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                if (!DrMk.Contains(" - "))
                                {
                                    Dr.SetUserProperty("comment", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;
                                }
                                else if (DrMk.Contains(" - "))
                                {
                                    for (int i = 1; i <= Xi; i++)
                                    {
                                        if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                        {
                                            int charIndex = i - 1;
                                            int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                            char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                            string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                            Dr.SetUserProperty("comment", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                            Dr_Sn++;
                                            charIndex += 1;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }

                        }
                        #endregion

                        txt_DrStartNumber.Text = Dr_Sn.ToString();
                    }  //Nếu chọn đánh mã kiểu ATAD 001  001A 001B
                    else
                    {
                        if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1) //nếu Write to không chọn, hoặc chọn Approved by
                        #region WRITE TO: Approved By
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 1)
                        #region Tittle 1
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.Title1 = Dr_Prefix + drStarNum + Dr_Posfix;
                                Dr.Modify();
                                Dr_Sn++;
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 2)
                        #region Tittle 2
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.Title2 = Dr_Prefix + drStarNum + Dr_Posfix;
                                Dr.Modify();
                                Dr_Sn++;
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 3)
                        #region Tittle 3
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.Title3 = Dr_Prefix + drStarNum + Dr_Posfix;
                                Dr.Modify();
                                Dr_Sn++;
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 4)
                        #region WRITE TO: User Field 1
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 5)
                        #region WRITE TO: User Field 2
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 6)
                        #region WRITE TO: User Field 3
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 7)
                        #region WRITE TO:Name
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.Name = Dr_Prefix + drStarNum + Dr_Posfix;
                                Dr.Modify();
                                Dr_Sn++;
                            }
                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 8)
                        #region WRITE TO: Checked By
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }

                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 9)
                        #region WRITE TO: Assigned To
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }

                        }
                        #endregion

                        if (comboBox1.SelectedIndex == 10)
                        #region WRITE TO: Comment
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                Dr.SetUserProperty("comment", Dr_Prefix + drStarNum + Dr_Posfix);
                                Dr_Sn++;
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }

                        }
                        #endregion

                        txt_DrStartNumber.Text = Dr_Sn.ToString();
                    }                                 //Nếu chọn đánh mã bình thường
                }
            }
            catch
            {
                MessageBox.Show("Wrong Tekla Structures Version or Start Number empty.");
            }
        }

        private void btn_NumberingDrawingPage_Click(object sender, EventArgs e)//button NumberingDrawingPage thông số bản vẽ (tab numbering)
        {
            try
            {
                tsd.DrawingEnumerator drawingList = drawingHandler.GetDrawingSelector().GetSelected();
                //Tao mot danh sach chua ban ve co kieu bien la loai moi duoc khoi tao DrawingMark
                List<DrawingMark> drawingMarks = new List<DrawingMark>();
                //Duyet danh sach drawingList de them vao danh sach DrawingMarks
                int numberingDigit = Convert.ToInt32(txt_Digit.Text);
                DialogResult dialogResult = MessageBox.Show(string.Format("Are you sure do you want to change attributes in {0} drawing?", drawingList.GetSize()), "Are you sure?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    foreach (tsd.Drawing dr in drawingList)
                    {
                        //Lay mark cua ban ve
                        string drMark = dr.Mark; // lấy được [B001] hoặc [B001 - 1]
                        string prefix = string.Empty;
                        //Split chuoi mark boi dau " - "
                        string[] subString = drMark.Split(new string[] { " - " }, StringSplitOptions.None); //bỏ " - " trong [B001 - 1], sẽ chia ra thành 2 "[B001" và "1]", việc này sẽ khiên gộp theo prefix bị lỗi, nên phải lấy assemblyPosition ở các bước tiếp theo đây.
                        int subIndex = 0; //chỉ số này để lấy các hậu tố 1, 2, 3... của drawing mark: ví dụ [B001 - 1] là 1, [B001 - 2] là 2,...

                        // các bước sau đây dể lấy assembly position vì dr.Mark có chứa '[' ']' nên lấy prefix bị sai
                        //Convert ban ve nay ve kieu AssemblyDrawing
                        tsd.AssemblyDrawing assemblyDrawing = dr as tsd.AssemblyDrawing;
                        //Thong qua ban ve ta co the lay assembly ngoai model
                        tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                        //Lay duoc mainPrefix
                        assembly.GetReportProperty("ASSEMBLY_POS", ref prefix);

                        if (subString.Length == 1) //trường hợp chỉ có 1 bản vẽ vd B001
                        {
                            // prefix = subString[0];
                            subIndex = 0;
                        }
                        else if (subString.Length == 2) //trường hợp có nhiều hơn bản vẽ vd B001 - 1, B001 - 2
                        {
                            //prefix = subString[0];
                            Match match = Regex.Match(subString[1], @"\d+");
                            subIndex = Convert.ToInt32(match.Value);
                        }
                        //MessageBox.Show(string.Format("{0}/{1}", prefix, subIndex.ToString())); //Kiểm tra
                        //Them vao vao danh sach Mark
                        DrawingMark drawingMark = new DrawingMark
                        {
                            Drawing = dr,
                            Prefix = prefix,
                            SubIndex = subIndex
                        };
                        drawingMarks.Add(drawingMark);
                    }

                    //Nhom danh sach theo mainpos
                    var group = drawingMarks.GroupBy(a => a.Prefix); //Gộp các bản vẽ có cùng Prefix vào 1 nhóm ( B001, B001 - 1, B001 - 2,...)

                    foreach (var item in group)
                    {

                        List<DrawingMark> drawingMarkList = item.Cast<DrawingMark>().ToList();//cast là 1 hình thức convert
                                                                                              //Sap xep drainglist theo subIndex
                        drawingMarkList = drawingMarkList.OrderBy(dr => dr.SubIndex).ToList(); // sắp xếp danh sách các DrawingMark theo chỉ sổ - 1, - 2, ...
                                                                                               //char[] alphabet = Enumerable.Range('A', 26).Select(x => (char)x).ToArray();
                        int index = Convert.ToInt32(txt_DrStartNumber.Text);

                        foreach (DrawingMark drawingMark in drawingMarkList)
                        {

                            //char curChar = alphabet[charIndex];
                            tsd.Drawing drawing = drawingMark.Drawing;
                            if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1)
                            {
                                drawing.SetUserProperty("DR_APPROVED_BY", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }
                            else if (comboBox1.SelectedIndex == 1)
                            {
                                drawing.Title1 = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                drawing.Modify();
                            }
                            else if (comboBox1.SelectedIndex == 2)
                            {
                                drawing.Title2 = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                drawing.Modify();
                            }
                            else if (comboBox1.SelectedIndex == 3)
                            {
                                drawing.Title3 = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                drawing.Modify();
                            }
                            else if (comboBox1.SelectedIndex == 4)
                            {
                                drawing.SetUserProperty("DRAWING_USERFIELD_1", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }
                            else if (comboBox1.SelectedIndex == 5)
                            {
                                drawing.SetUserProperty("DRAWING_USERFIELD_2", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }
                            else if (comboBox1.SelectedIndex == 6)
                            {
                                drawing.SetUserProperty("DRAWING_USERFIELD_3", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }
                            else if (comboBox1.SelectedIndex == 7)
                            {
                                drawing.Name = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                drawing.Modify();
                            }
                            else if (comboBox1.SelectedIndex == 8)
                            {
                                drawing.SetUserProperty("DR_CHECKED_BY", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }
                            else if (comboBox1.SelectedIndex == 9)
                            {
                                drawing.SetUserProperty("DR_ASSIGNED_TO", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }
                            else if (comboBox1.SelectedIndex == 10)
                            {
                                drawing.SetUserProperty("comment", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                            }

                            index++;
                        }
                    }
                    Operation.RunMacro("UDA_A&M_Modify.cs");
                }
            }
            catch { }
        }

        private void btn_ResetStartNo_Click(object sender, EventArgs e)//button NumberingDrawingPage thông số bản vẽ (tab numbering)
        {
            txt_DrStartNumber.Text = "1";
        }

        private void chb_AtadNumberingStyle_CheckedChanged(object sender, EventArgs e)
        {
            if (!chb_AtadNumberingStyle.Checked)
            {
                chb_AtadNumberingStyle.Text = "001, 002, 003, 004, 005, 006, 007...";
            }
            else
            {
                chb_AtadNumberingStyle.Text = "001, 001A, 001B, 002, 002A, 002B, 003,...";
            }
        }


        private void btn_RunNumbering_Click(object sender, EventArgs e)
        {
            if (cb_NumberingChangeNumberingpage.SelectedIndex == -1 || cb_NumberingChangeNumberingpage.SelectedIndex == 0)
            {
                try
                {
                    tsd.Drawing MyDrawing = drawingHandler.GetActiveDrawing();
                    string Dr_Prefix = txt_DrPrefix.Text; // lấy dữ liệu từ ô Prefix.
                    string Dr_Posfix = txt_DrPostfix.Text; // lấy dữ liệu từ ô Postfix.
                    int Dr_Sn = Convert.ToInt32(txt_DrStartNumber.Text); // lấy dữ liệu từ ô start no.
                    int numberingDigit = Convert.ToInt32(txt_Digit.Text);
                    tsd.DrawingEnumerator MyDrEnum = drawingHandler.GetDrawingSelector().GetSelected();
                    int Xi = MyDrEnum.GetSize();
                    //MessageBox.Show(Xi.ToString());
                    char[] alphabet = Enumerable.Range('A', 26).Select(x => (char)x).ToArray();
                    tsd.Drawing Dr = null;
                    DialogResult dialogResult = MessageBox.Show(string.Format("Are you sure do you want to change attributes in {0} drawing?", Xi), "Are you sure?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        if (chb_AtadNumberingStyle.Checked)
                        {
                            if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1) //nếu Write to không chọn, hoặc chọn Approved by
                            #region WRITE TO: Approved By
                            {
                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }
                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 1)
                            #region Tittle 1
                            {
                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.Title1 = Dr_Prefix + drStarNum + Dr_Posfix;
                                        Dr.Modify();
                                        Dr_Sn++;
                                    }

                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.Title1 = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                                Dr.Modify();
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 2)
                            #region Tittle 2
                            {
                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.Title2 = Dr_Prefix + drStarNum + Dr_Posfix;
                                        Dr.Modify();
                                        Dr_Sn++;
                                    }

                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.Title2 = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                                Dr.Modify();
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 3)
                            #region Tittle 3
                            {
                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.Title3 = Dr_Prefix + drStarNum + Dr_Posfix;
                                        Dr.Modify();
                                        Dr_Sn++;
                                    }

                                    else if (DrMk.Contains(" - "))
                                    {

                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.Title3 = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                                Dr.Modify();
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 4)
                            #region WRITE TO: User Field 1
                            {
                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 5)
                            #region WRITE TO: User Field 2
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 6)
                            #region WRITE TO: User Field 3
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 7)
                            #region WRITE TO:Name
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.Name = Dr_Prefix + drStarNum + Dr_Posfix;
                                        Dr.Modify();
                                        Dr_Sn++;
                                    }

                                    else if (DrMk.Contains(" - "))
                                    {

                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.Name = Dr_Prefix + drStarNum1 + curChar + Dr_Posfix;
                                                Dr.Modify();
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 8)
                            #region WRITE TO: Checked By
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }

                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 9)
                            #region WRITE TO: Assigned To
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }

                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 10)
                            #region WRITE TO: Comment
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    if (!DrMk.Contains(" - "))
                                    {
                                        Dr.SetUserProperty("comment", Dr_Prefix + drStarNum + Dr_Posfix);
                                        Dr_Sn++;
                                    }
                                    else if (DrMk.Contains(" - "))
                                    {
                                        for (int i = 1; i <= Xi; i++)
                                        {
                                            if (DrMk.Contains(" - " + i.ToString())) //NẾU DRAWING MARK DẠNG B001 - 1, B001 - 2, ...
                                            {
                                                int charIndex = i - 1;
                                                int Dr_Sn1 = Dr_Sn -= 1; // SỐ THỨ TỰ BẰNG SỐ THỨ TỰ BẢN VẼ B001 (SAU KHI ĐÁNH SỐ B001 ĐÃ TĂNG 1 NÊN PHẢI -1)
                                                char curChar = alphabet[charIndex]; //LẤY KÍ TỰ ALPHABET THỨ charIndex;
                                                string drStarNum1 = Dr_Sn1.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                                Dr.SetUserProperty("comment", Dr_Prefix + drStarNum1 + curChar + Dr_Posfix);
                                                Dr_Sn++;
                                                charIndex += 1;
                                                break;
                                            }
                                        }
                                    }


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }

                            }
                            #endregion

                            txt_DrStartNumber.Text = Dr_Sn.ToString();
                        }  //Nếu chọn đánh mã kiểu ATAD 001  001A 001B
                        else
                        {
                            if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1) //nếu Write to không chọn, hoặc chọn Approved by
                            #region WRITE TO: Approved By
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 1)
                            #region Tittle 1
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.Title1 = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;


                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 2)
                            #region Tittle 2
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.Title2 = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;


                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 3)
                            #region Tittle 3
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.Title3 = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;


                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 4)
                            #region WRITE TO: User Field 1
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 5)
                            #region WRITE TO: User Field 2
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 6)
                            #region WRITE TO: User Field 3
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 7)
                            #region WRITE TO:Name
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.Name = Dr_Prefix + drStarNum + Dr_Posfix;
                                    Dr.Modify();
                                    Dr_Sn++;


                                }
                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 8)
                            #region WRITE TO: Checked By
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }

                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 9)
                            #region WRITE TO: Assigned To
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }

                            }
                            #endregion

                            if (comboBox1.SelectedIndex == 10)
                            #region WRITE TO: Comment
                            {


                                while (MyDrEnum.MoveNext())
                                {
                                    Dr = MyDrEnum.Current as Drawing;
                                    string DrMk = Dr.Mark;
                                    string drStarNum = Dr_Sn.ToString().PadLeft(numberingDigit, '0'); // CHUYỂN SỐ THỬ TỰ BẢN VẼ VỀ DẠNG 001 HOẶC 01 HOẶC 00001...
                                    Dr.SetUserProperty("comment", Dr_Prefix + drStarNum + Dr_Posfix);
                                    Dr_Sn++;


                                }
                                if (Dr is GADrawing)
                                {
                                    Operation.RunMacro("UDA_G_Modify.cs");
                                }
                                else
                                {
                                    Operation.RunMacro("UDA_A&M_Modify.cs");
                                }

                            }
                            #endregion

                            txt_DrStartNumber.Text = Dr_Sn.ToString();
                        }                                 //Nếu chọn đánh mã bình thường
                    }
                }
                catch
                {
                    MessageBox.Show("Wrong Tekla Structures Version or Start Number empty.");
                }
            }
            else if (cb_NumberingChangeNumberingpage.SelectedIndex == 1)
            {
                try
                {
                    tsd.Drawing MyDrawing = drawingHandler.GetActiveDrawing();
                    string Dr_Prefix = txt_DrPrefix.Text;
                    string Dr_Posfix = txt_DrPostfix.Text;
                    tsd.DrawingEnumerator MyDrEnum = drawingHandler.GetDrawingSelector().GetSelected();
                    int Xi = MyDrEnum.GetSize();
                    //MessageBox.Show(Xi.ToString());
                    tsd.Drawing Dr = null;
                    DialogResult dialogResult = MessageBox.Show(string.Format("Are you sure do you want to change attributes in {0} drawing?", Xi), "Are you sure?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1)
                        #region WRITE TO: Approved By
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("DR_APPROVED_BY", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 1)
                        #region Tittle 1
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                Dr.Title1 = Dr_Prefix + Dr_Posfix;
                                Dr.Modify();
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 2)
                        #region Tittle 2
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                Dr.Title2 = Dr_Prefix + Dr_Posfix;
                                Dr.Modify();
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 3)
                        #region Tittle 3
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                Dr.Title3 = Dr_Prefix + Dr_Posfix;
                                Dr.Modify();
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 4)
                        #region WRITE TO: User Field 1
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("DRAWING_USERFIELD_1", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 5)
                        #region WRITE TO: User Field 2
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("DRAWING_USERFIELD_2", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 6)
                        #region WRITE TO: User Field 3
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("DRAWING_USERFIELD_3", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 7)
                        #region WRITE TO: Name
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                Dr.Name = Dr_Prefix + Dr_Posfix;
                                Dr.Modify();
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 8)
                        #region WRITE TO: Checked By
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("DR_CHECKED_BY", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 9)
                        #region WRITE TO: ASSIGNED TO
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("DR_ASSIGNED_TO", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion

                        else if (comboBox1.SelectedIndex == 10)
                        #region WRITE TO: Comment
                        {
                            while (MyDrEnum.MoveNext())
                            {
                                Dr = MyDrEnum.Current as Drawing;
                                string DrMk = Dr.Mark;
                                Dr.SetUserProperty("comment", Dr_Prefix + Dr_Posfix);
                            }
                            if (Dr is GADrawing)
                            {
                                Operation.RunMacro("UDA_G_Modify.cs");
                            }
                            else
                            {
                                Operation.RunMacro("UDA_A&M_Modify.cs");
                            }
                            MessageBox.Show("Changing complete -- Done");
                        }
                        #endregion
                    }
                    else
                    { }
                }
                catch
                {
                    MessageBox.Show("Wrong Tekla Structures Version or Start Number empty.");
                }
            }
            else if (cb_NumberingChangeNumberingpage.SelectedIndex == 2)
            {
                try
                {
                    tsd.DrawingEnumerator drawingList = drawingHandler.GetDrawingSelector().GetSelected();
                    //Tao mot danh sach chua ban ve co kieu bien la loai moi duoc khoi tao DrawingMark
                    List<DrawingMark> drawingMarks = new List<DrawingMark>();
                    //Duyet danh sach drawingList de them vao danh sach DrawingMarks
                    int numberingDigit = Convert.ToInt32(txt_Digit.Text);

                    DialogResult dialogResult = MessageBox.Show(string.Format("Are you sure do you want to change attributes in {0} drawing?", drawingList.GetSize()), "Are you sure?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        foreach (tsd.Drawing dr in drawingList)
                        {
                            //Lay mark cua ban ve
                            string drMark = dr.Mark; // lấy được [B001] hoặc [B001 - 1]
                            string prefix = string.Empty;
                            //Split chuoi mark boi dau " - "
                            string[] subString = drMark.Split(new string[] { " - " }, StringSplitOptions.None); //bỏ " - " trong [B001 - 1], sẽ chia ra thành 2 "[B001" và "1]", việc này sẽ khiên gộp theo prefix bị lỗi, nên phải lấy assemblyPosition ở các bước tiếp theo đây.
                            int subIndex = 0; //chỉ số này để lấy các hậu tố 1, 2, 3... của drawing mark: ví dụ [B001 - 1] là 1, [B001 - 2] là 2,...

                            // các bước sau đây dể lấy assembly position vì dr.Mark có chứa '[' ']' nên lấy prefix bị sai
                            //Convert ban ve nay ve kieu AssemblyDrawing
                            tsd.AssemblyDrawing assemblyDrawing = dr as tsd.AssemblyDrawing;
                            //Thong qua ban ve ta co the lay assembly ngoai model
                            tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                            //Lay duoc mainPrefix
                            assembly.GetReportProperty("ASSEMBLY_POS", ref prefix);

                            if (subString.Length == 1) //trường hợp chỉ có 1 bản vẽ vd B001
                            {
                                // prefix = subString[0];
                                subIndex = 0;
                            }
                            else if (subString.Length == 2) //trường hợp có nhiều hơn bản vẽ vd B001 - 1, B001 - 2
                            {
                                //prefix = subString[0];
                                Match match = Regex.Match(subString[1], @"\d+");
                                subIndex = Convert.ToInt32(match.Value);
                            }
                            //MessageBox.Show(string.Format("{0}/{1}", prefix, subIndex.ToString())); //Kiểm tra
                            //Them vao vao danh sach Mark
                            DrawingMark drawingMark = new DrawingMark
                            {
                                Drawing = dr,
                                Prefix = prefix,
                                SubIndex = subIndex
                            };
                            drawingMarks.Add(drawingMark);
                        }

                        //Nhom danh sach theo mainpos
                        var group = drawingMarks.GroupBy(a => a.Prefix); //Gộp các bản vẽ có cùng Prefix vào 1 nhóm ( B001, B001 - 1, B001 - 2,...)

                        foreach (var item in group)
                        {

                            List<DrawingMark> drawingMarkList = item.Cast<DrawingMark>().ToList();//cast là 1 hình thức convert
                                                                                                  //Sap xep drainglist theo subIndex
                            drawingMarkList = drawingMarkList.OrderBy(dr => dr.SubIndex).ToList(); // sắp xếp danh sách các DrawingMark theo chỉ sổ - 1, - 2, ...
                                                                                                   //char[] alphabet = Enumerable.Range('A', 26).Select(x => (char)x).ToArray();
                            int index = Convert.ToInt32(txt_DrStartNumber.Text);

                            foreach (DrawingMark drawingMark in drawingMarkList)
                            {

                                //char curChar = alphabet[charIndex];
                                tsd.Drawing drawing = drawingMark.Drawing;
                                if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == -1)
                                {
                                    drawing.SetUserProperty("DR_APPROVED_BY", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }
                                else if (comboBox1.SelectedIndex == 1)
                                {
                                    drawing.Title1 = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                    drawing.Modify();
                                }
                                else if (comboBox1.SelectedIndex == 2)
                                {
                                    drawing.Title2 = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                    drawing.Modify();
                                }
                                else if (comboBox1.SelectedIndex == 3)
                                {
                                    drawing.Title3 = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                    drawing.Modify();
                                }
                                else if (comboBox1.SelectedIndex == 4)
                                {
                                    drawing.SetUserProperty("DRAWING_USERFIELD_1", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }
                                else if (comboBox1.SelectedIndex == 5)
                                {
                                    drawing.SetUserProperty("DRAWING_USERFIELD_2", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }
                                else if (comboBox1.SelectedIndex == 6)
                                {
                                    drawing.SetUserProperty("DRAWING_USERFIELD_3", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }
                                else if (comboBox1.SelectedIndex == 7)
                                {
                                    drawing.Name = string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0'));
                                    drawing.Modify();
                                }
                                else if (comboBox1.SelectedIndex == 8)
                                {
                                    drawing.SetUserProperty("DR_CHECKED_BY", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }
                                else if (comboBox1.SelectedIndex == 9)
                                {
                                    drawing.SetUserProperty("DR_ASSIGNED_TO", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }
                                else if (comboBox1.SelectedIndex == 10)
                                {
                                    drawing.SetUserProperty("comment", string.Format("{0}/{1}", index.ToString().PadLeft(numberingDigit, '0'), drawingMarkList.Count.ToString().PadLeft(numberingDigit, '0')));
                                }

                                index++;


                            }
                        }
                        Operation.RunMacro("UDA_A&M_Modify.cs");
                    }
                }
                catch { }
            }
        }


        private void btn_OpenActiveDrawing_Click(object sender, EventArgs e)//button OpenActiveDrawing mở bản vẽ đã ban hành (Export)
        {
            try
            {
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                string[] directoryFull = null; //đường dẫn để mở file
                string filePath = string.Empty; //đường dẫn trong ô người dùng nhập
                string modelPath = model.GetInfo().ModelPath;
                string drAsMark = string.Empty; // Tên bản vẽ Assembly
                string drSgMark = string.Empty; // Tên bản vẽ Single-part
                string lastRevDr = string.Empty;

                // các bước sau đây dể lấy assembly position vì dr.Mark có chứa '[' ']' nên lấy actived_dr.Mark bị sai
                if (actived_dr is tsd.AssemblyDrawing)
                {
                    tsd.AssemblyDrawing assemblyDrawing = actived_dr as tsd.AssemblyDrawing;
                    //Thong qua ban ve ta co the lay assembly ngoai model
                    tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                    //Lay duoc mainPrefix
                    assembly.GetReportProperty("ASSEMBLY_POS", ref drAsMark);
                    if (txt_RevisionDrawing.Text == "")
                        assembly.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                    else
                        lastRevDr = txt_RevisionDrawing.Text;

                    if (tbx_DirectoryOpenDrawing.Text == "")
                        filePath = modelPath;
                    else
                    {
                        filePath = tbx_DirectoryOpenDrawing.Text;
                        if (filePath.Last() != '\\')
                            filePath = tbx_DirectoryOpenDrawing.Text + '\\';
                    }
                    if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", drAsMark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                    else
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", drAsMark, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Mark và revision
                                                                                                                                                      // MessageBox.Show(directory[0]);
                    Process.Start(directoryFull[0]);//mở file đầu tiên trong array list directoryFull
                }
                else if (actived_dr is tsd.SinglePartDrawing)
                {
                    tsd.SinglePartDrawing singleDrawing = actived_dr as tsd.SinglePartDrawing;
                    tsm.Part part = model.SelectModelObject(singleDrawing.PartIdentifier) as tsm.Part;
                    part.GetReportProperty("PART_POS", ref drSgMark);
                    drSgMark = part.GetPartMark();
                    if (txt_RevisionDrawing.Text == "")
                        part.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                    else
                        lastRevDr = txt_RevisionDrawing.Text;
                    if (tbx_DirectoryOpenDrawing.Text == "")
                        filePath = modelPath;
                    else
                    {
                        filePath = tbx_DirectoryOpenDrawing.Text;
                        if (filePath.Last() != '\\')
                            filePath = tbx_DirectoryOpenDrawing.Text + '\\';
                    }
                    if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", drSgMark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                    else
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", drSgMark, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Mark và revision
                                                                                                                                                      // MessageBox.Show(directory[0]);
                    Process.Start(directoryFull[0]);//mở file đầu tiên trong array list directoryFull
                }
                else if (actived_dr is tsd.GADrawing)
                {
                    tsd.GADrawing gADrawing = actived_dr as tsd.GADrawing;
                    if (tbx_DirectoryOpenDrawing.Text == "")
                        filePath = modelPath;
                    else
                    {
                        filePath = tbx_DirectoryOpenDrawing.Text;
                        if (filePath.Last() != '\\')
                            filePath = tbx_DirectoryOpenDrawing.Text + '\\';
                    }
                    String gaTitle1 = gADrawing.Title1;
                    lastRevDr = txt_RevisionDrawing.Text;

                    if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", gaTitle1, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Title1 và revision
                    else
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", gaTitle1, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Title1 và revision
                                                                                                                                                      // MessageBox.Show(directory[0]);
                    Process.Start(directoryFull[0]);//mở file đầu tiên trong array list directoryFull
                }
            }
            catch
            {
                MessageBox.Show("Bản vẽ không tìm thấy, thử nhập số rev bản vẽ vào ô Rev rồi mở lại.");
            }
        }

        private void btn_DimSectionSelectView_Click(object sender, EventArgs e)
        {
            tsd.AssemblyDrawing assemblyDrawing = drawingHandler.GetActiveDrawing() as tsd.AssemblyDrawing; //Lấy bản vẽ assembly
            tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
            tsm.Part mainPart = assembly.GetMainPart() as tsm.Part;
            string mainPartProfileType = string.Empty;
            mainPart.GetReportProperty("PROFILE_TYPE", ref mainPartProfileType);

            if (cb_DimSectionSelect.SelectedIndex == 0 || cb_DimSectionSelect.SelectedIndex == -1) // Chọn Select Views (nếu chọn 1 là  Select Parts)
            {
                try
                {
                    tsd.DrawingEnumerator.AutoFetch = true;
                    tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                    tsd.DrawingObjectEnumerator drObjEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                    foreach (tsd.View curView in drObjEnum)
                    {
                        tsd.View view = curView as tsd.View;
                        tsd.DrawingObjectEnumerator drobjenum = view.GetAllObjects(new Type[] { typeof(tsd.Part) }); //
                        PartDistribution partDistribution = new PartDistribution(mainPart, drobjenum); //để phân loại Danh sách các parts có trong view -> phân loại mainpart, secondaries
                        tsd.ViewBase viewBase = partDistribution.ViewBase;
                        // tsd.ViewBase viewBase = drobjenum.Current.GetView() as tsd.ViewBase;
                        //MessageBox.Show(drobjenum.GetSize().ToString());

                        tsm.Part topFlangeSection = partDistribution.TopFlangeSection;
                        tsm.Part botFlangeSection = partDistribution.BotFlangeSection;
                        List<tsm.Part> partsTamThayBoltTron = partDistribution.ListTamThayBoltTron;
                        List<tsm.Part> partsTamDung = partDistribution.ListTamDung;
                        List<tsm.Part> partsTamNgang = partDistribution.ListTamNgang;
                        List<tsm.Part> partsNghiengPhai = partDistribution.ListTamNghiengPhai;
                        List<tsm.Part> partsNghiengTrai = partDistribution.ListTamNghiengTrai;

                        tsd.PointList listPointBoltDim_X = new PointList();
                        tsd.PointList listPointBoltDim_Y = new PointList();

                        Part_Edge partEdgemainPart = new Part_Edge(view, mainPart);
                        t3d.Point mainPartPointMidTop = partEdgemainPart.PointMidTop;
                        t3d.Point mainPartPointMidBot = partEdgemainPart.PointMidBot;
                        t3d.Point mainPartPointXminYmin = partEdgemainPart.PointXminYmin;
                        t3d.Point mainPartPointXminYmax = partEdgemainPart.PointXminYmax;
                        t3d.Point mainPartPointXmaxYmax = partEdgemainPart.PointXmaxYmax;
                        t3d.Point mainPartPointXmaxYmin = partEdgemainPart.PointXmaxYmin;
                        tsd.PointList pointListDimXTotal = new PointList(); // ds điểm để dim phương X tấm thấy bolt tròn. Chỉ lấy các tấm nằm bên trái và phải mainpart.
                        if (Control.ModifierKeys == Keys.Shift)
                        {

                        }
                        else
                        {
                            pointListDimXTotal.Add(mainPartPointXminYmax);
                            pointListDimXTotal.Add(mainPartPointXmaxYmax);
                        }
                        // MessageBox.Show(partsTamThayBoltTron.Count().ToString()) ;
                        if (mainPartProfileType == "B") // DIM CHO THEP TO HOP
                        {
                            try
                            {
                                Part_Edge partEdgeBotFlangeSection = new Part_Edge(view, botFlangeSection); //đoạn này và đoạn dưới đổi chỗ cho nhau sẽ bị lỗi check lại****
                                Part_Edge partEdgeTopFlangeSection = new Part_Edge(view, topFlangeSection);

                                if (partEdgeTopFlangeSection.List_Edge != null && partEdgeBotFlangeSection.List_Edge == null) //Nếu cánh trên không lỗi, cánh dưới lỗi, thì lấy cánh trên đo (cho point cánh dưới = cánh trên)
                                {
                                    mainPartPointMidTop = partEdgeTopFlangeSection.PointMidTop;
                                    mainPartPointXminYmax = partEdgeTopFlangeSection.PointXminYmax;
                                    mainPartPointXmaxYmax = partEdgeTopFlangeSection.PointXmaxYmax;
                                    mainPartPointMidBot = mainPartPointMidTop;
                                    mainPartPointXminYmin = mainPartPointXminYmax;
                                    mainPartPointXmaxYmin = mainPartPointXmaxYmax;
                                }
                                else if (partEdgeBotFlangeSection.List_Edge != null && partEdgeTopFlangeSection.List_Edge == null) //Nếu cánh dưới không lỗi, cánh trên lỗi, thì lấy cánh dưới đo (cho point cánh trên = cánh dưới)
                                {
                                    mainPartPointMidBot = partEdgeBotFlangeSection.PointMidBot;
                                    mainPartPointXminYmin = partEdgeBotFlangeSection.PointXminYmin;
                                    mainPartPointXmaxYmin = partEdgeBotFlangeSection.PointXmaxYmin;
                                    mainPartPointMidTop = mainPartPointMidBot;
                                    mainPartPointXminYmax = mainPartPointXminYmin;
                                    mainPartPointXmaxYmax = mainPartPointXmaxYmin;
                                }
                                else if (partEdgeBotFlangeSection.List_Edge != null && partEdgeTopFlangeSection.List_Edge != null)// trường hợp cả 2 cánh trên dưới đều ok
                                {
                                    mainPartPointMidTop = partEdgeTopFlangeSection.PointMidTop;
                                    mainPartPointXminYmax = partEdgeTopFlangeSection.PointXminYmax;
                                    mainPartPointXmaxYmax = partEdgeTopFlangeSection.PointXmaxYmax;
                                    mainPartPointMidBot = partEdgeBotFlangeSection.PointMidBot;
                                    mainPartPointXminYmin = partEdgeBotFlangeSection.PointXminYmin;
                                    mainPartPointXmaxYmin = partEdgeBotFlangeSection.PointXmaxYmin;
                                }
                            }
                            catch { }
                            //MessageBox.Show(mainPart.GetPartMark());
                            #region Duyet tat cả các part cua  partsTamThayBoltTron

                            foreach (tsm.Part tamThayBoltTron in partsTamThayBoltTron) //Duyet tat cả các part cua  partsTamThayBoltTron
                            {
                                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                tsd.PointList pointListDimY2 = new PointList(); //ds điểm dim bolt max Y và điểm part edge có Y gần bolt max Y nhất (dim bolt ra cạnh gần nhất của part phương Y)
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge partEdgeTamThayBoltTron = new Part_Edge(view, tamThayBoltTron); // ds chưa tất cả điểm part edge cua  partEdgeTamThayBoltTron
                                List<t3d.Point> listpartEdgeTamThayBoltTron = partEdgeTamThayBoltTron.List_Edge;//ds điểm part edge cua tamThayBoltTron
                                t3d.Point pointTamThayBoltTronXminYmin = partEdgeTamThayBoltTron.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = partEdgeTamThayBoltTron.PointXmaxYmax; // góc trên bên phải
                                t3d.Point pointTamThayBoltTronXminYmax = partEdgeTamThayBoltTron.PointXminYmax; // góc trên bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmin = partEdgeTamThayBoltTron.PointXmaxYmin; // góc dưới bên phải
                                if (pointTamThayBoltTronXminYmin.X > mainPartPointMidTop.X) //Nằm bên PHẢI mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimXTotal.Add(p);
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                            pointListBolt.Add(p);
                                        }
                                    }

                                    t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                                    List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên PHẢI
                                    foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                                    {
                                        if (point.Y > pointBoltMaxY.Y && point.X > pointBoltMaxY.X)
                                        {
                                            points2.Add(point);
                                        }
                                    }
                                    t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                                    pointListDimY2.Add(point1);
                                    pointListDimY2.Add(pointBoltMaxY);
                                    //MessageBox.Show("TamThayBoltTron ben phai");
                                    if (Control.ModifierKeys == Keys.Shift)
                                    {
                                        pointListDimXTotal.Add(pointTamThayBoltTronXminYmax);
                                        pointListDimY.Add(pointTamThayBoltTronXmaxYmax);
                                        pointListDimY.Add(pointTamThayBoltTronXmaxYmin);

                                    }
                                    else
                                    {
                                        pointListDimX.Add(mainPartPointMidTop);
                                        pointListDimY.Add(mainPartPointXmaxYmax);
                                    }
                                    if (chbox_DimClose.Checked)
                                    {
                                        pointListDimY.Add(mainPartPointXmaxYmin);
                                    }
                                    //createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 200); //tạo dim phương X
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY, cb_DimentionAttributes.Text, 200); //tạo dim phương Y
                                    if (chbox_DimInternalPart.Checked)
                                    {
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 120 - Convert.ToInt32(Math.Abs(pointListDimY[0].X - pointListDimY2[0].X))); //tạo dim phương Y cho bolt ra mép cao nhất
                                    }
                                }//Nằm bên PHAI mid point mainpart

                                else if (pointTamThayBoltTronXmaxYmax.X < mainPartPointMidTop.X) //Nằm bên TRÁI mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimXTotal.Add(p);
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                            pointListBolt.Add(p);
                                        }
                                    }
                                    t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                                    List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên TRAI
                                    foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                                    {
                                        if (point.Y > pointBoltMaxY.Y && point.X < pointBoltMaxY.X)
                                        {
                                            points2.Add(point);
                                        }
                                    }
                                    t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                                    pointListDimY2.Add(point1);
                                    pointListDimY2.Add(pointBoltMaxY);

                                    if (Control.ModifierKeys == Keys.Shift)
                                    {
                                        pointListDimXTotal.Add(pointTamThayBoltTronXmaxYmax);
                                        pointListDimY.Add(pointTamThayBoltTronXminYmax);
                                        pointListDimY.Add(pointTamThayBoltTronXminYmin);

                                    }
                                    else
                                    {
                                        pointListDimX.Add(mainPartPointMidTop);
                                        pointListDimY.Add(mainPartPointXminYmax);
                                    }
                                    if (chbox_DimClose.Checked)
                                    {
                                        pointListDimY.Add(mainPartPointXminYmin); // điểm góc dưới bên trái mainpart
                                    }
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    //createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 200);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked)
                                    {
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 120 - Convert.ToInt32(Math.Abs(pointListDimY[0].X - pointListDimY2[0].X))); //tạo dim phương Y cho bolt ra mép cao nhất
                                    }


                                }//Nằm bên TRÁI mid point mainpart

                                else if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y) //Nằm bên TRÊN mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                        }
                                    }
                                    pointListDimX.Add(mainPartPointMidTop);
                                    pointListDimY.Add(mainPartPointXminYmax);
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    try
                                    {
                                        createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                    }
                                    catch  //Tấm plate có 1 hàng  bolt dọc nằm giữa
                                    {
                                        pointListDimX.Add(mainPartPointXminYmax);
                                        pointListDimX.Add(mainPartPointXmaxYmax);
                                        createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    }

                                }//Nằm bên TRÊN mid point mainpart

                                else if (pointTamThayBoltTronXminYmin.Y < mainPartPointMidBot.Y) //Nằm bên DƯỚI mid point mainpart 
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                        }
                                    }
                                    pointListDimX.Add(mainPartPointMidBot);
                                    pointListDimY.Add(mainPartPointXminYmin);
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                } //Nằm bên DƯỚI mid point mainpart 
                            } //Duyet tat cả các part cua  partsTamThayBoltTron
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTotal, cb_DimentionAttributes.Text, 250);
                            #endregion

                            #region Duyet tat cả các part cua  partsTamDung
                            foreach (tsm.Part tamDung in partsTamDung) //Duyet tat cả các part cua  partsTamDung
                            {
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeTamDung = new Part_Edge(view, tamDung); // ds chưa tất cả điểm part edge cua  tamDung
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                                t3d.Point pointTamThayBoltTronXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                                t3d.Point pointTamThayBoltTronXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải
                                t3d.Point pointTamThayBoltTronXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                                if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y)
                                {
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax);// add thêm điểm góc trên bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax);// add thêm điểm góc trên bên phải
                                    }
                                    //Đoạn này dim Bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_Y.Clear();
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                        foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_Y.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75); //tạo dim Bolt
                                    }

                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);

                                }   //tấm đứng nằm PHÍA TRÊN
                                else if (pointTamThayBoltTronXmaxYmin.Y < mainPartPointMidBot.Y)
                                {
                                    pointListDimX.Add(mainPartPointXminYmin);
                                    pointListDimX.Add(mainPartPointXmaxYmin);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin); // add thêm điểm góc dưới bên trái
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin); // add thêm điểm góc dưới bên trái
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    //Đoạn này dim bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_Y.Clear();
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                        foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_Y.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);

                                }   //tấm đứng nằm PHÍA DƯỚI
                                else if (pointTamThayBoltTronXminYmax.X >= mainPartPointXmaxYmax.X && pointTamThayBoltTronXmaxYmax.X > mainPartPointXmaxYmax.X)
                                // điểm nhỏ nhất của EndPlate phương X chênh lệch với điểm lớn nhất phương X của main part k quá 5mm và điểm lớn nhất của nó phải lớn hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_Y.Clear();
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                    pointListDimX.Add(pointTamThayBoltTronXmaxYmax);
                                    pointListDimY.Add(pointTamThayBoltTronXmaxYmax);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương Y
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmax);
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmin);
                                    foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_Y.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm đứng nằm Bên Phải
                                else if (pointTamThayBoltTronXmaxYmax.X <= mainPartPointXminYmax.X && pointTamThayBoltTronXminYmax.X < mainPartPointXminYmax.X)
                                // điểm nhỏ nhất của EndPlate phương X chênh lệch với điểm lớn nhất phương X của main part k quá 5mm và điểm lớn nhất của nó phải lớn hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_Y.Clear();
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXminYmax);
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXminYmin);
                                    pointListDimX.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimY.Add(pointTamThayBoltTronXminYmax);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương Y-
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmin);
                                    foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_Y.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm đứng nằm Bên Trái
                            } //Duyet tat cả các part cua  partsTamDung
                            #endregion

                            #region Duyet tat cả các part cua  partsTamNgang
                            foreach (tsm.Part tamNgang in partsTamNgang) //Duyet tat cả các part cua  partsTamNgang
                            {
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeTamDung = new Part_Edge(view, tamNgang); // ds chưa tất cả điểm part edge cua  tamDung
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                                t3d.Point pointTamThayBoltTronXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                                t3d.Point pointTamThayBoltTronXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải
                                t3d.Point pointTamThayBoltTronXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                                if (pointTamThayBoltTronXmaxYmax.X < mainPartPointMidTop.X) //tấm Ngang nằm bên TRÁI
                                {
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimX.Add(mainPartPointXminYmin);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    //Đoạn này dim Bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                        foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);
                                }//tấm Ngang nằm bên TRÁI
                                else if (pointTamThayBoltTronXminYmax.X > mainPartPointMidBot.X) //tấm Ngang nằm bên Phải
                                {
                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    pointListDimX.Add(mainPartPointXmaxYmin);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax); // add thêm điểm góc trên bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax); // add thêm điểm góc trên bên phải
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    //Đoạn này dim bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                        foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);
                                }//tấm Ngang nằm bên Phải
                                else if (pointTamThayBoltTronXminYmin.Y >= mainPartPointXminYmax.Y && pointTamThayBoltTronXmaxYmax.Y > mainPartPointXminYmax.Y)
                                // điểm nhỏ nhất của EndPlate phương X chênh lệch với điểm lớn nhất phương X của main part k quá 5mm và điểm lớn nhất của nó phải lớn hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_X.Clear();
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                    pointListDimX.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimY.Add(pointTamThayBoltTronXminYmax);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương X
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmax);
                                    foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_X.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm ngang nằm trên
                                else if (pointTamThayBoltTronXminYmax.Y <= mainPartPointXminYmin.Y && pointTamThayBoltTronXmaxYmin.Y < mainPartPointXminYmin.Y)
                                // điểm lớn nhất của EndPlate phương Y chênh lệch với điểm nhỏ nhất phương Y của main part k quá 5mm và điểm nhỏ nhất của nó phải nhỏ hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_X.Clear();
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXminYmin);
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmin);
                                    pointListDimX.Add(pointTamThayBoltTronXminYmin);
                                    pointListDimY.Add(pointTamThayBoltTronXminYmin);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương X
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmin);
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmin);
                                    foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_X.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm ngang nằm Dưới
                            } //Duyet tat cả các part cua  partsTamNgang
                            #endregion

                            #region Duyet tat cả các part cua  partsNghiengPhai
                            foreach (var tamNghiengPhai in partsNghiengPhai)//Duyet tat cả các part cua  partsNghiengPhai
                            {
                                tsd.PointList pointListDimXTamNghieng = new PointList(); // ds điểm để dim phương X cho tam nghieng
                                tsd.PointList pointListDimYTamNghieng = new PointList(); // ds điểm để dim phương Y cho tam nghieng
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeNghiengPhai = new Part_Edge(view, tamNghiengPhai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengPhai.List_Edge;//ds điểm của part 
                                t3d.Point tamNghiengPhaiXmin = part_EdgeNghiengPhai.PointXmin0;
                                t3d.Point tamNghiengPhaiXmax = part_EdgeNghiengPhai.PointXmax0;
                                t3d.Point tamNghiengPhaiYmin = part_EdgeNghiengPhai.PointYmin0;
                                t3d.Point tamNghiengPhaiYmax = part_EdgeNghiengPhai.PointYmax0;

                                if (Math.Abs(tamNghiengPhaiYmin.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengPhai nằm bên TRÊN cây H
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiYmin);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengPhaiYmin);
                                    pointListPartDimX.Add(tamNghiengPhaiXmax);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 200);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmin);
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmax);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmin, tamNghiengPhaiXmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }//tấm NghiengPhai nằm bên TRÊN cây H
                                else if (Math.Abs(tamNghiengPhaiYmax.Y - mainPartPointMidBot.Y) < 3) //tấm NghiengPhai nằm bên DƯỚI cây H
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiYmax);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengPhaiYmax);
                                    pointListPartDimX.Add(tamNghiengPhaiXmin);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmax);
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmin, tamNghiengPhaiXmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }//tấm NghiengPhai nằm bên DƯỚI cây H
                                else if (tamNghiengPhaiXmax.X < mainPartPointMidBot.X && tamNghiengPhaiXmax.Y > mainPartPointMidBot.Y && tamNghiengPhaiXmax.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên TRÁI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiXmax);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimY.Add(tamNghiengPhaiXmax);
                                    pointListPartDimY.Add(tamNghiengPhaiYmin);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmax);
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiXmax, tamNghiengPhaiYmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                                else if (tamNghiengPhaiXmin.X > mainPartPointMidBot.X && tamNghiengPhaiXmin.Y > mainPartPointMidBot.Y && tamNghiengPhaiXmin.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên PHẢI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimY.Add(tamNghiengPhaiYmax);
                                    pointListPartDimY.Add(tamNghiengPhaiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmax);
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmax, tamNghiengPhaiXmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }

                            }//Duyet tat cả các part cua  partsNghiengPhai
                            #endregion

                            #region Duyet tat cả các part cua  partsNghiengTrai
                            foreach (var tamNghiengTrai in partsNghiengTrai)//Duyet tat cả các part cua  partsNghiengTrai
                            {
                                tsd.PointList pointListDimXTamNghieng = new PointList(); // ds điểm để dim phương X cho tam nghieng
                                tsd.PointList pointListDimYTamNghieng = new PointList(); // ds điểm để dim phương Y cho tam nghieng
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeNghiengTrai = new Part_Edge(view, tamNghiengTrai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengTrai.List_Edge;//ds điểm của part 
                                t3d.Point tamNghiengTraiXmin = part_EdgeNghiengTrai.PointXmin0;
                                t3d.Point tamNghiengTraiXmax = part_EdgeNghiengTrai.PointXmax0;
                                t3d.Point tamNghiengTraiYmin = part_EdgeNghiengTrai.PointYmin0;
                                t3d.Point tamNghiengTraiYmax = part_EdgeNghiengTrai.PointYmax0;

                                if (Math.Abs(tamNghiengTraiYmin.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengTrai nằm bên TRÊN
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiYmin);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengTraiXmin);
                                    pointListPartDimX.Add(tamNghiengTraiYmin);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmin);
                                        listPointBoltDim_X.Add(tamNghiengTraiYmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmin, tamNghiengTraiYmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }

                                }
                                else if (Math.Abs(tamNghiengTraiYmax.Y - mainPartPointMidBot.Y) < 3) //tấm NghiengPhai nằm bên DƯỚI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiYmax);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengTraiXmax);
                                    pointListPartDimX.Add(tamNghiengTraiYmax);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                        listPointBoltDim_X.Add(tamNghiengTraiYmax);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiYmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                                else if (tamNghiengTraiXmax.X < mainPartPointMidBot.X && tamNghiengTraiXmax.Y > mainPartPointMidBot.Y && tamNghiengTraiXmax.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên TRÁI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiXmax);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương Y
                                    pointListPartDimY.Add(tamNghiengTraiYmax);
                                    pointListPartDimY.Add(tamNghiengTraiXmax);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                        listPointBoltDim_X.Add(tamNghiengTraiYmax);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiYmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                                else if (tamNghiengTraiXmin.X > mainPartPointMidBot.X && tamNghiengTraiXmin.Y > mainPartPointMidBot.Y && tamNghiengTraiXmin.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên PHẢI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương Y
                                    pointListPartDimY.Add(tamNghiengTraiXmax);
                                    pointListPartDimY.Add(tamNghiengTraiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                        listPointBoltDim_X.Add(tamNghiengTraiXmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiXmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                            }//Duyet tat cả các part cua  partsNghiengTrai
                            #endregion

                            Operation.DisplayPrompt("Mainpart Mark: " + mainPart.GetPartMark());

                            //MessageBox.Show(topFlangeSection.GetPartMark());
                            //MessageBox.Show(botFlangeSection.GetPartMark());
                        }
                        else if (mainPartProfileType == "I")
                        {
                            #region Duyet tat cả các part cua  partsTamThayBoltTron
                            foreach (tsm.Part tamThayBoltTron in partsTamThayBoltTron) //Duyet tat cả các part cua  partsTamThayBoltTron
                            {
                                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                tsd.PointList pointListDimY2 = new PointList(); //ds điểm dim bolt max Y và điểm part edge có Y gần bolt max Y nhất (dim bolt ra cạnh gần nhất của part phương Y)
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge partEdgeTamThayBoltTron = new Part_Edge(view, tamThayBoltTron); // ds chưa tất cả điểm part edge cua  partEdgeTamThayBoltTron
                                List<t3d.Point> listpartEdgeTamThayBoltTron = partEdgeTamThayBoltTron.List_Edge;//ds điểm part edge cua tamThayBoltTron
                                t3d.Point pointTamThayBoltTronXminYmin = partEdgeTamThayBoltTron.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = partEdgeTamThayBoltTron.PointXmaxYmax; // góc trên bên phải

                                if (pointTamThayBoltTronXminYmin.X > mainPartPointMidTop.X) //Nằm bên PHẢI mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                            pointListBolt.Add(p);
                                        }
                                    }

                                    t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                                    List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên PHẢI
                                    foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                                    {
                                        if (point.Y > pointBoltMaxY.Y && point.X > pointBoltMaxY.X)
                                        {
                                            points2.Add(point);
                                        }
                                    }
                                    t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                                    pointListDimY2.Add(point1);
                                    pointListDimY2.Add(pointBoltMaxY);
                                    //MessageBox.Show("TamThayBoltTron ben phai");
                                    pointListDimX.Add(mainPartPointMidTop);
                                    pointListDimY.Add(mainPartPointXmaxYmax);
                                    if (chbox_DimClose.Checked)
                                    {
                                        pointListDimY.Add(mainPartPointXmaxYmin);
                                    }
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 200); //tạo dim phương X
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY, cb_DimentionAttributes.Text, 200); //tạo dim phương Y
                                    if (chbox_DimInternalPart.Checked)
                                    {
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 120 - Convert.ToInt32(Math.Abs(pointListDimY[0].X - pointListDimY2[0].X))); //tạo dim phương Y cho bolt ra mép cao nhất
                                    }

                                }//Nằm bên PHAI mid point mainpart

                                else if (pointTamThayBoltTronXmaxYmax.X < mainPartPointMidTop.X) //Nằm bên TRÁI mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                            pointListBolt.Add(p);
                                        }
                                    }
                                    t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                                    List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên TRAI
                                    foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                                    {
                                        if (point.Y > pointBoltMaxY.Y && point.X < pointBoltMaxY.X)
                                        {
                                            points2.Add(point);
                                        }
                                    }
                                    t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                                    pointListDimY2.Add(point1);
                                    pointListDimY2.Add(pointBoltMaxY);

                                    pointListDimX.Add(mainPartPointMidTop);
                                    pointListDimY.Add(mainPartPointXminYmax); // điểm góc Trên bên trái mainpart
                                    if (chbox_DimClose.Checked)
                                    {
                                        pointListDimY.Add(mainPartPointXminYmin); // điểm góc dưới bên trái mainpart
                                    }
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 200);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked)
                                    {
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 120 - Convert.ToInt32(Math.Abs(pointListDimY[0].X - pointListDimY2[0].X))); //tạo dim phương Y cho bolt ra mép cao nhất
                                    }


                                }//Nằm bên TRÁI mid point mainpart

                                else if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y) //Nằm bên TRÊN mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                        }
                                    }
                                    pointListDimX.Add(mainPartPointMidTop);
                                    pointListDimY.Add(mainPartPointXminYmax);
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    try
                                    {
                                        createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                    }
                                    catch  //Tấm plate có 1 hàng  bolt dọc nằm giữa
                                    {
                                        pointListDimX.Add(mainPartPointXminYmax);
                                        pointListDimX.Add(mainPartPointXmaxYmax);
                                        createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    }

                                }//Nằm bên TRÊN mid point mainpart

                                else if (pointTamThayBoltTronXminYmin.Y < mainPartPointMidBot.Y) //Nằm bên DƯỚI mid point mainpart 
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                        }
                                    }
                                    pointListDimX.Add(mainPartPointMidBot);
                                    pointListDimY.Add(mainPartPointXminYmin);
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                } //Nằm bên DƯỚI mid point mainpart 
                            } //Duyet tat cả các part cua  partsTamThayBoltTron
                            #endregion

                            #region Duyet tat cả các part cua  partsTamDung
                            foreach (tsm.Part tamDung in partsTamDung) //Duyet tat cả các part cua  partsTamDung
                            {
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeTamDung = new Part_Edge(view, tamDung); // ds chưa tất cả điểm part edge cua  tamDung
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                                t3d.Point pointTamThayBoltTronXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                                t3d.Point pointTamThayBoltTronXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải
                                t3d.Point pointTamThayBoltTronXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                                if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y)
                                {
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax);// add thêm điểm góc trên bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax);// add thêm điểm góc trên bên phải
                                    }
                                    //Đoạn này dim Bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_Y.Clear();
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                        foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_Y.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75); //tạo dim Bolt
                                    }

                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);

                                }   //tấm đứng nằm PHÍA TRÊN
                                else if (pointTamThayBoltTronXmaxYmin.Y < mainPartPointMidBot.Y)
                                {
                                    pointListDimX.Add(mainPartPointXminYmin);
                                    pointListDimX.Add(mainPartPointXmaxYmin);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin); // add thêm điểm góc dưới bên trái
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin); // add thêm điểm góc dưới bên trái
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    //Đoạn này dim bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_Y.Clear();
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                        listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                        foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_Y.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);

                                }   //tấm đứng nằm PHÍA DƯỚI
                                else if (pointTamThayBoltTronXminYmax.X >= mainPartPointXmaxYmax.X && pointTamThayBoltTronXmaxYmax.X > mainPartPointXmaxYmax.X)
                                // điểm nhỏ nhất của EndPlate phương X chênh lệch với điểm lớn nhất phương X của main part k quá 5mm và điểm lớn nhất của nó phải lớn hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_Y.Clear();
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                    pointListDimX.Add(pointTamThayBoltTronXmaxYmax);
                                    pointListDimY.Add(pointTamThayBoltTronXmaxYmax);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương Y
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmax);
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmin);
                                    foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_Y.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm đứng nằm Bên Phải
                                else if (pointTamThayBoltTronXmaxYmax.X <= mainPartPointXminYmax.X && pointTamThayBoltTronXminYmax.X < mainPartPointXminYmax.X)
                                // điểm nhỏ nhất của EndPlate phương X chênh lệch với điểm lớn nhất phương X của main part k quá 5mm và điểm lớn nhất của nó phải lớn hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_Y.Clear();
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXminYmax);
                                    listPointBoltDim_Y.Add(pointTamThayBoltTronXminYmin);
                                    pointListDimX.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimY.Add(pointTamThayBoltTronXminYmax);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương Y-
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmin);
                                    foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_Y.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm đứng nằm Bên Trái
                            } //Duyet tat cả các part cua  partsTamDung
                            #endregion

                            #region Duyet tat cả các part cua  partsTamNgang
                            foreach (tsm.Part tamNgang in partsTamNgang) //Duyet tat cả các part cua  partsTamNgang
                            {
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeTamDung = new Part_Edge(view, tamNgang); // ds chưa tất cả điểm part edge cua  tamDung
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                                t3d.Point pointTamThayBoltTronXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                                t3d.Point pointTamThayBoltTronXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải
                                t3d.Point pointTamThayBoltTronXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                                if (pointTamThayBoltTronXmaxYmax.X < mainPartPointMidTop.X) //tấm Ngang nằm bên TRÁI
                                {
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimX.Add(mainPartPointXminYmin);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                        pointListDimX.Add(pointTamThayBoltTronXminYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    //Đoạn này dim Bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                        foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);
                                }//tấm Ngang nằm bên TRÁI
                                else if (pointTamThayBoltTronXminYmax.X > mainPartPointMidBot.X) //tấm Ngang nằm bên Phải
                                {
                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    pointListDimX.Add(mainPartPointXmaxYmin);
                                    if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax); // add thêm điểm góc trên bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                                    {
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmax); // add thêm điểm góc trên bên phải
                                        pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                                    }
                                    //Đoạn này dim bolt vào part
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                        listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                        foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);
                                }//tấm Ngang nằm bên Phải
                                else if (pointTamThayBoltTronXminYmin.Y >= mainPartPointXminYmax.Y && pointTamThayBoltTronXmaxYmax.Y > mainPartPointXminYmax.Y)
                                // điểm nhỏ nhất của EndPlate phương X chênh lệch với điểm lớn nhất phương X của main part k quá 5mm và điểm lớn nhất của nó phải lớn hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_X.Clear();
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                    pointListDimX.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimY.Add(pointTamThayBoltTronXminYmax);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương X
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmax);
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmax);
                                    foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_X.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm ngang nằm trên
                                else if (pointTamThayBoltTronXminYmax.Y <= mainPartPointXminYmin.Y && pointTamThayBoltTronXmaxYmin.Y < mainPartPointXminYmin.Y)
                                // điểm lớn nhất của EndPlate phương Y chênh lệch với điểm nhỏ nhất phương Y của main part k quá 5mm và điểm nhỏ nhất của nó phải nhỏ hơn mainpart
                                {
                                    //Đoạn này dim bolt vào part
                                    listPointBoltDim_X.Clear();
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXminYmin);
                                    listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmin);
                                    pointListDimX.Add(pointTamThayBoltTronXminYmin);
                                    pointListDimY.Add(pointTamThayBoltTronXminYmin);
                                    tsd.PointList pointListDimPart = new PointList(); // ds điểm để dim phương X
                                    pointListDimPart.Add(pointTamThayBoltTronXminYmin);
                                    pointListDimPart.Add(pointTamThayBoltTronXmaxYmin);
                                    foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            listPointBoltDim_X.Add(p);
                                        }
                                    }
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimPart, cb_DimentionAttributes.Text, 120);//tạo dim Part
                                }   //tấm ngang nằm Dưới
                            } //Duyet tat cả các part cua  partsTamNgang
                            #endregion

                            #region Duyet tat cả các part cua  partsNghiengPhai
                            foreach (var tamNghiengPhai in partsNghiengPhai)//Duyet tat cả các part cua  partsNghiengPhai
                            {
                                tsd.PointList pointListDimXTamNghieng = new PointList(); // ds điểm để dim phương X cho tam nghieng
                                tsd.PointList pointListDimYTamNghieng = new PointList(); // ds điểm để dim phương Y cho tam nghieng
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeNghiengPhai = new Part_Edge(view, tamNghiengPhai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengPhai.List_Edge;//ds điểm của part 
                                t3d.Point tamNghiengPhaiXmin = part_EdgeNghiengPhai.PointXmin0;
                                t3d.Point tamNghiengPhaiXmax = part_EdgeNghiengPhai.PointXmax0;
                                t3d.Point tamNghiengPhaiYmin = part_EdgeNghiengPhai.PointYmin0;
                                t3d.Point tamNghiengPhaiYmax = part_EdgeNghiengPhai.PointYmax0;

                                if (Math.Abs(tamNghiengPhaiYmin.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengPhai nằm bên TRÊN cây H
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiYmin);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengPhaiYmin);
                                    pointListPartDimX.Add(tamNghiengPhaiXmax);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 200);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmin);
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmax);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmin, tamNghiengPhaiXmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }//tấm NghiengPhai nằm bên TRÊN cây H
                                else if (Math.Abs(tamNghiengPhaiYmax.Y - mainPartPointMidBot.Y) < 3) //tấm NghiengPhai nằm bên DƯỚI cây H
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiYmax);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengPhaiYmax);
                                    pointListPartDimX.Add(tamNghiengPhaiXmin);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmax);
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmin, tamNghiengPhaiXmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }//tấm NghiengPhai nằm bên DƯỚI cây H
                                else if (tamNghiengPhaiXmax.X < mainPartPointMidBot.X && tamNghiengPhaiXmax.Y > mainPartPointMidBot.Y && tamNghiengPhaiXmax.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên TRÁI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiXmax);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimY.Add(tamNghiengPhaiXmax);
                                    pointListPartDimY.Add(tamNghiengPhaiYmin);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmax);
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiXmax, tamNghiengPhaiYmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                                else if (tamNghiengPhaiXmin.X > mainPartPointMidBot.X && tamNghiengPhaiXmin.Y > mainPartPointMidBot.Y && tamNghiengPhaiXmin.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên PHẢI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengPhaiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimY.Add(tamNghiengPhaiYmax);
                                    pointListPartDimY.Add(tamNghiengPhaiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 200);
                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengPhaiYmax);
                                        listPointBoltDim_X.Add(tamNghiengPhaiXmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmax, tamNghiengPhaiXmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }

                            }//Duyet tat cả các part cua  partsNghiengPhai
                            #endregion

                            #region Duyet tat cả các part cua  partsNghiengTrai
                            foreach (var tamNghiengTrai in partsNghiengTrai)//Duyet tat cả các part cua  partsNghiengTrai
                            {
                                tsd.PointList pointListDimXTamNghieng = new PointList(); // ds điểm để dim phương X cho tam nghieng
                                tsd.PointList pointListDimYTamNghieng = new PointList(); // ds điểm để dim phương Y cho tam nghieng
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge part_EdgeNghiengTrai = new Part_Edge(view, tamNghiengTrai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                                List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengTrai.List_Edge;//ds điểm của part 
                                t3d.Point tamNghiengTraiXmin = part_EdgeNghiengTrai.PointXmin0;
                                t3d.Point tamNghiengTraiXmax = part_EdgeNghiengTrai.PointXmax0;
                                t3d.Point tamNghiengTraiYmin = part_EdgeNghiengTrai.PointYmin0;
                                t3d.Point tamNghiengTraiYmax = part_EdgeNghiengTrai.PointYmax0;

                                if (Math.Abs(tamNghiengTraiYmin.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengTrai nằm bên TRÊN
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiYmin);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengTraiXmin);
                                    pointListPartDimX.Add(tamNghiengTraiYmin);
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmin);
                                        listPointBoltDim_X.Add(tamNghiengTraiYmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmin, tamNghiengTraiYmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }

                                }
                                else if (Math.Abs(tamNghiengTraiYmax.Y - mainPartPointMidBot.Y) < 3) //tấm NghiengPhai nằm bên DƯỚI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiYmax);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                                    pointListPartDimX.Add(tamNghiengTraiXmax);
                                    pointListPartDimX.Add(tamNghiengTraiYmax);
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                        listPointBoltDim_X.Add(tamNghiengTraiYmax);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiYmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                                else if (tamNghiengTraiXmax.X < mainPartPointMidBot.X && tamNghiengTraiXmax.Y > mainPartPointMidBot.Y && tamNghiengTraiXmax.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên TRÁI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXminYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiXmax);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương Y
                                    pointListPartDimY.Add(tamNghiengTraiYmax);
                                    pointListPartDimY.Add(tamNghiengTraiXmax);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                        listPointBoltDim_X.Add(tamNghiengTraiYmax);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiYmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                                else if (tamNghiengTraiXmin.X > mainPartPointMidBot.X && tamNghiengTraiXmin.Y > mainPartPointMidBot.Y && tamNghiengTraiXmin.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên PHẢI
                                {
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmax);
                                    pointListDimXTamNghieng.Add(mainPartPointXmaxYmin);
                                    pointListDimXTamNghieng.Add(tamNghiengTraiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                                    //Đoạn dưới để xác định góc nghiêng và dim bolt
                                    tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương Y
                                    pointListPartDimY.Add(tamNghiengTraiXmax);
                                    pointListPartDimY.Add(tamNghiengTraiXmin);
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 150);

                                    if (chbox_DimInternalPart.Checked == true)
                                    {
                                        listPointBoltDim_X.Clear();
                                        listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                        listPointBoltDim_X.Add(tamNghiengTraiXmin);
                                        foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                        {
                                            foreach (t3d.Point p in bgr.BoltPositions)
                                            {
                                                listPointBoltDim_X.Add(p);
                                            }
                                        }
                                        createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiXmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                                    }
                                }
                            }//Duyet tat cả các part cua  partsNghiengTrai
                            #endregion
                        }
                        else if (mainPartProfileType == "U" || mainPartProfileType == "L")
                        {
                            //dim cho ListTamThayBoltTron
                            foreach (tsm.Part tamThayBoltTron in partsTamThayBoltTron) //Duyet tat cả các part cua  partsTamThayBoltTron
                            {
                                tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                                tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                                tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                                tsd.PointList pointListDimY2 = new PointList(); //ds điểm dim bolt max Y và điểm part edge có Y gần bolt max Y nhất (dim bolt ra cạnh gần nhất của part phương Y)
                                List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                                Part_Edge partEdgeTamThayBoltTron = new Part_Edge(view, tamThayBoltTron); // ds chưa tất cả điểm part edge cua  partEdgeTamThayBoltTron
                                List<t3d.Point> listpartEdgeTamThayBoltTron = partEdgeTamThayBoltTron.List_Edge;//ds điểm part edge cua tamThayBoltTron
                                t3d.Point pointTamThayBoltTronXminYmin = partEdgeTamThayBoltTron.PointXminYmin; // góc dưới bên traí
                                t3d.Point pointTamThayBoltTronXmaxYmax = partEdgeTamThayBoltTron.PointXmaxYmax; // góc trên bên phải

                                if (pointTamThayBoltTronXmaxYmax.X > mainPartPointXmaxYmax.X) //Nằm bên PHẢI point X max mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                            pointListBolt.Add(p);
                                        }
                                    }
                                    t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                                    List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên PHẢI
                                    foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                                    {
                                        if (point.Y > pointBoltMaxY.Y && point.X > pointBoltMaxY.X)
                                        {
                                            points2.Add(point);
                                        }
                                    }
                                    t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                                    pointListDimY2.Add(point1);
                                    pointListDimY2.Add(pointBoltMaxY);
                                    //MessageBox.Show("TamThayBoltTron ben phai");

                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimY.Add(mainPartPointXmaxYmax);
                                    if (chbox_DimClose.Checked)
                                    {
                                        pointListDimY.Add(mainPartPointXmaxYmin);
                                    }
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250); //tạo dim phương X
                                    createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250); //tạo dim phương Y
                                    if (chbox_DimInternalPart.Checked)
                                    {
                                        createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 150); //tạo dim phương Y cho bolt ra mép cao nhất
                                    }

                                }//Nằm bên TRÁI mid point mainpart

                                else if (pointTamThayBoltTronXminYmin.X < mainPartPointMidTop.X) //Nằm bên TRÁI mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                            pointListBolt.Add(p);
                                        }
                                    }

                                    t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                                    List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên PHẢI
                                    foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                                    {
                                        if (point.Y > pointBoltMaxY.Y && point.X < pointBoltMaxY.X)
                                        {
                                            points2.Add(point);
                                        }
                                    }
                                    t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                                    pointListDimY2.Add(point1);
                                    pointListDimY2.Add(pointBoltMaxY);

                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimY.Add(mainPartPointXminYmax);
                                    if (chbox_DimClose.Checked)
                                    {
                                        pointListDimY.Add(mainPartPointXminYmin);
                                    }
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                    if (chbox_DimInternalPart.Checked)
                                    {
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 150); //tạo dim phương Y cho bolt ra mép cao nhất
                                    }

                                }//Nằm bên TRÁI mid point mainpart

                                else if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y) //Nằm bên TRÊN mid point mainpart
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                        }
                                    }
                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimY.Add(mainPartPointXminYmax);
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    try
                                    {
                                        createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                        createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                    }
                                    catch  //Tấm plate có 1 hàng  bolt dọc nằm giữa
                                    {
                                        pointListDimX.Add(mainPartPointXminYmax);
                                        pointListDimX.Add(mainPartPointXmaxYmax);
                                        createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    }

                                }//Nằm bên TRÊN mid point mainpart

                                else if (pointTamThayBoltTronXminYmin.Y < mainPartPointMidBot.Y) //Nằm bên DƯỚI mid point mainpart 
                                {
                                    foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                                    {
                                        foreach (t3d.Point p in bgr.BoltPositions)
                                        {
                                            pointListDimX.Add(p);
                                            pointListDimY.Add(p);
                                        }
                                    }
                                    pointListDimX.Add(mainPartPointXmaxYmax);
                                    pointListDimX.Add(mainPartPointXminYmax);
                                    pointListDimY.Add(mainPartPointXminYmin);
                                    //MessageBox.Show("TamThayBoltTron ben Trai");
                                    createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                    createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                                } //Nằm bên DƯỚI mid point mainpart 
                            } //Duyet tat cả các part cua  partsTamThayBoltTron
                        }
                        Add_Mark();
                    }
                }
                catch (Exception)
                { }
            }
            else if (cb_DimSectionSelect.SelectedIndex == 1)
            {
                try
                {
                    tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                    //tsm.TransformationPlane curplane = wph.GetCurrentTransformationPlane(); // lay toa do hien hanh
                    //wph.SetCurrentTransformationPlane(new TransformationPlane());
                    tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                    CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                    PartDistribution partDistribution = new PartDistribution(mainPart, drobjenum);
                    tsd.View view = partDistribution.View;
                    tsd.ViewBase viewBase = partDistribution.ViewBase;

                    //tsm.Part mainPart = partDistribution.MainPart;
                    List<tsm.Part> partsTamThayBoltTron = partDistribution.ListTamThayBoltTron;
                    List<tsm.Part> partsTamDung = partDistribution.ListTamDung;
                    List<tsm.Part> partsTamNgang = partDistribution.ListTamNgang;
                    List<tsm.Part> partsNghiengPhai = partDistribution.ListTamNghiengPhai;
                    List<tsm.Part> partsNghiengTrai = partDistribution.ListTamNghiengTrai;

                    tsd.PointList listPointBoltDim_X = new PointList();
                    tsd.PointList listPointBoltDim_Y = new PointList();

                    tsdui.Picker c_picker = drawingHandler.GetPicker();
                    var pick = c_picker.PickPoint("Chon diem thu nhat");
                    //Part_Edge partEdgemainPart = new Part_Edge(view, mainPart);
                    t3d.Point mainPartPointMidTop = pick.Item1;

                    //MessageBox.Show(mainPartPointMidTop.X.ToString() + mainPartPointMidTop.Y.ToString());
                    //dim cho ListTamThayBoltTron
                    foreach (tsm.Part tamThayBoltTron in partsTamThayBoltTron) //Duyet tat cả các part cua  partsTamThayBoltTron
                    {
                        tsd.StraightDimensionSetHandler straightDimensionSetHandler = new StraightDimensionSetHandler();
                        tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                        tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                        tsd.PointList pointListDimY2 = new PointList(); //ds điểm dim bolt max Y và điểm part edge có Y gần bolt max Y nhất (dim bolt ra cạnh gần nhất của part phương Y)
                        List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                        Part_Edge partEdgeTamThayBoltTron = new Part_Edge(view, tamThayBoltTron); // ds chưa tất cả điểm part edge cua  partEdgeTamThayBoltTron
                        List<t3d.Point> listpartEdgeTamThayBoltTron = partEdgeTamThayBoltTron.List_Edge;//ds điểm part edge cua tamThayBoltTron
                        t3d.Point pointTamThayBoltTronXminYmin = partEdgeTamThayBoltTron.PointXminYmin; // góc dưới bên traí
                        t3d.Point pointTamThayBoltTronXmaxYmax = partEdgeTamThayBoltTron.PointXmaxYmax; // góc trên bên phải

                        if (pointTamThayBoltTronXminYmin.X > mainPartPointMidTop.X) //Nằm bên PHẢI mid point mainpart
                        {
                            foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    pointListDimX.Add(p);
                                    pointListDimY.Add(p);
                                    pointListBolt.Add(p);
                                }
                            }

                            t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                            List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên PHẢI
                            foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                            {
                                if (point.Y > pointBoltMaxY.Y && point.X > pointBoltMaxY.X)
                                {
                                    points2.Add(point);
                                }
                            }
                            t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                            pointListDimY2.Add(point1);
                            pointListDimY2.Add(pointBoltMaxY);
                            //MessageBox.Show("TamThayBoltTron ben phai");
                            pointListDimX.Add(mainPartPointMidTop);
                            pointListDimY.Add(mainPartPointMidTop);
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 200); //tạo dim phương X
                            createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY, cb_DimentionAttributes.Text, 200); //tạo dim phương Y
                            if (chbox_DimInternalPart.Checked)
                            {
                                createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 150); //tạo dim phương Y cho bolt ra mép cao nhất
                            }

                        }//Nằm bên TRÁI mid point mainpart

                        else if (pointTamThayBoltTronXmaxYmax.X < mainPartPointMidTop.X) //Nằm bên TRÁI mid point mainpart
                        {
                            foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    pointListDimX.Add(p);
                                    pointListDimY.Add(p);
                                    pointListBolt.Add(p);
                                }
                            }
                            t3d.Point pointBoltMaxY = pointListBolt.OrderByDescending(a => a.Y).ToList()[0]; //điểm bolt có Y cao nhất
                            List<t3d.Point> points2 = new List<t3d.Point>();//ds điểm của part nằm bên PHẢI
                            foreach (t3d.Point point in listpartEdgeTamThayBoltTron) //duyệt các điểm part egde để lấy điểm có tọa độ Y lớn hơn điểm boltmax Y
                            {
                                if (point.Y > pointBoltMaxY.Y && point.X < pointBoltMaxY.X)
                                {
                                    points2.Add(point);
                                }
                            }
                            t3d.Point point1 = points2.OrderBy(a => a.Y).ToList()[0]; //Điểm Y của part gần bolt max Y nhất.
                            pointListDimY2.Add(point1);
                            pointListDimY2.Add(pointBoltMaxY);

                            pointListDimX.Add(mainPartPointMidTop);
                            pointListDimY.Add(mainPartPointMidTop);
                            //MessageBox.Show("TamThayBoltTron ben Trai");
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 200);
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 200);
                            if (chbox_DimInternalPart.Checked)
                            {
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY2, cb_DimentionAttributes.Text, 150); //tạo dim phương Y cho bolt ra mép cao nhất
                            }

                        }//Nằm bên TRÁI mid point mainpart

                        else if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y) //Nằm bên TRÊN mid point mainpart
                        {
                            foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    pointListDimX.Add(p);
                                    pointListDimY.Add(p);
                                }
                            }
                            pointListDimX.Add(mainPartPointMidTop);
                            pointListDimY.Add(mainPartPointMidTop);
                            //MessageBox.Show("TamThayBoltTron ben Trai");
                            try
                            {
                                createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                                createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                            }
                            catch  //Tấm plate có 1 hàng  bolt dọc nằm giữa
                            {
                                createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                            }

                        }//Nằm bên TRÊN mid point mainpart

                        else if (pointTamThayBoltTronXminYmin.Y < mainPartPointMidTop.Y) //Nằm bên DƯỚI mid point mainpart 
                        {
                            foreach (tsm.BoltGroup bgr in tamThayBoltTron.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    pointListDimX.Add(p);
                                    pointListDimY.Add(p);
                                }
                            }
                            pointListDimX.Add(mainPartPointMidTop);
                            pointListDimY.Add(mainPartPointMidTop);
                            //MessageBox.Show("TamThayBoltTron ben Trai");
                            createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 250);
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimY, cb_DimentionAttributes.Text, 250);
                        } //Nằm bên DƯỚI mid point mainpart 
                    } //Duyet tat cả các part cua  partsTamThayBoltTron

                    foreach (tsm.Part tamDung in partsTamDung) //Duyet tat cả các part cua  partsTamDung
                    {
                        tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                        tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                        List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                        Part_Edge part_EdgeTamDung = new Part_Edge(view, tamDung); // ds chưa tất cả điểm part edge cua  tamDung
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                        t3d.Point pointTamThayBoltTronXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                        t3d.Point pointTamThayBoltTronXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                        t3d.Point pointTamThayBoltTronXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải
                        t3d.Point pointTamThayBoltTronXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                        if (pointTamThayBoltTronXmaxYmax.Y > mainPartPointMidTop.Y) //tấm đứng nằm PHÍA TRÊN
                        {
                            pointListDimX.Add(mainPartPointMidTop);
                            if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                            {
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmax);// add thêm điểm góc trên bên phải
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmax);// add thêm điểm góc trên bên phải
                            }
                            //Đoạn này dim Bolt vào part
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_Y.Clear();
                                listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_Y.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75); //tạo dim Bolt
                            }

                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);

                        }
                        else if (pointTamThayBoltTronXmaxYmin.Y < mainPartPointMidTop.Y) //tấm đứng nằm PHÍA DƯỚI
                        {
                            pointListDimX.Add(mainPartPointMidTop);
                            if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmin); // add thêm điểm góc dưới bên trái
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                            {
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmin); // add thêm điểm góc dưới bên trái
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                            }
                            //Đoạn này dim bolt vào part
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_Y.Clear();
                                listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmax);
                                listPointBoltDim_Y.Add(pointTamThayBoltTronXmaxYmin);
                                foreach (tsm.BoltGroup bgr in tamDung.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_Y.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_Y(viewBase, listPointBoltDim_Y, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                            createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);

                        }
                    } //Duyet tat cả các part cua  partsTamDung

                    foreach (tsm.Part tamNgang in partsTamNgang) //Duyet tat cả các part cua  partsTamNgang
                    {
                        tsd.PointList pointListDimX = new PointList(); // ds điểm để dim phương X
                        tsd.PointList pointListDimY = new PointList(); // ds điểm để dim phương Y
                        List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                        Part_Edge part_EdgeTamDung = new Part_Edge(view, tamNgang); // ds chưa tất cả điểm part edge cua  tamDung
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeTamDung.List_Edge;//ds điểm của part 
                        t3d.Point pointTamThayBoltTronXminYmin = part_EdgeTamDung.PointXminYmin; // góc dưới bên traí
                        t3d.Point pointTamThayBoltTronXmaxYmax = part_EdgeTamDung.PointXmaxYmax; // góc trên bên phải
                        t3d.Point pointTamThayBoltTronXmaxYmin = part_EdgeTamDung.PointXmaxYmin; // góc dưới bên phải
                        t3d.Point pointTamThayBoltTronXminYmax = part_EdgeTamDung.PointXminYmax; // góc trên bên trái
                        if (pointTamThayBoltTronXmaxYmax.X < mainPartPointMidTop.X) //tấm Ngang nằm bên TRÁI
                        {
                            pointListDimX.Add(mainPartPointMidTop);
                            if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmin);// add thêm điểm góc dưới bên phải
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                            {
                                pointListDimX.Add(pointTamThayBoltTronXminYmax); // add thêm điểm góc trên bên trái
                                pointListDimX.Add(pointTamThayBoltTronXminYmin);// add thêm điểm góc dưới bên phải
                            }
                            //Đoạn này dim Bolt vào part
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);
                        }//tấm Ngang nằm bên TRÁI
                        else if (pointTamThayBoltTronXminYmax.X > mainPartPointMidTop.X) //tấm Ngang nằm bên Phải
                        {
                            pointListDimX.Add(mainPartPointMidTop);
                            if (imcb_DimensionTo.SelectedIndex == 0 || imcb_DimensionTo.SelectedIndex == -1) // dim to LEFT TOP
                            {
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmax); // add thêm điểm góc trên bên phải
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 1) // dim to RIGHT BOTTOM
                            {
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                            }
                            else if (imcb_DimensionTo.SelectedIndex == 2) // dim to BOTH SIDE
                            {
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmax); // add thêm điểm góc trên bên phải
                                pointListDimX.Add(pointTamThayBoltTronXmaxYmin);// add thêm điểm góc dưới bên phải
                            }
                            //Đoạn này dim bolt vào part
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(pointTamThayBoltTronXminYmax);
                                listPointBoltDim_X.Add(pointTamThayBoltTronXmaxYmax);
                                foreach (tsm.BoltGroup bgr in tamNgang.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_X(viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                            createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimX, cb_DimentionAttributes.Text, 450);
                        }//tấm Ngang nằm bên Phải
                    } //Duyet tat cả các part cua  partsTamNgang

                    foreach (var tamNghiengPhai in partsNghiengPhai)//Duyet tat cả các part cua  partsNghiengPhai
                    {
                        tsd.PointList pointListDimXTamNghieng = new PointList(); // ds điểm để dim phương X cho tam nghieng
                        tsd.PointList pointListDimYTamNghieng = new PointList(); // ds điểm để dim phương Y cho tam nghieng
                        List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                        Part_Edge part_EdgeNghiengPhai = new Part_Edge(view, tamNghiengPhai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengPhai.List_Edge;//ds điểm của part 
                        t3d.Point tamNghiengPhaiXmin = part_EdgeNghiengPhai.PointXmin0;
                        t3d.Point tamNghiengPhaiXmax = part_EdgeNghiengPhai.PointXmax0;
                        t3d.Point tamNghiengPhaiYmin = part_EdgeNghiengPhai.PointYmin0;
                        t3d.Point tamNghiengPhaiYmax = part_EdgeNghiengPhai.PointYmax0;

                        if (Math.Abs(tamNghiengPhaiYmin.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengPhai nằm bên TRÊN cây H
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengPhaiYmin);
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                            pointListPartDimX.Add(tamNghiengPhaiYmin);
                            pointListPartDimX.Add(tamNghiengPhaiXmax);
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 200);

                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengPhaiYmin);
                                listPointBoltDim_X.Add(tamNghiengPhaiXmax);
                                foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmin, tamNghiengPhaiXmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }
                        else if (Math.Abs(tamNghiengPhaiYmax.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengPhai nằm bên DƯỚI cây H
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengPhaiYmax);
                            createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                            pointListPartDimX.Add(tamNghiengPhaiYmax);
                            pointListPartDimX.Add(tamNghiengPhaiXmin);
                            createDimension.CreateStraightDimensionSet_X_(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 200);
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengPhaiYmax);
                                listPointBoltDim_X.Add(tamNghiengPhaiXmin);
                                foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmin, tamNghiengPhaiXmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }
                        else if (tamNghiengPhaiXmax.X < mainPartPointMidTop.X && tamNghiengPhaiXmax.Y > mainPartPointMidTop.Y && tamNghiengPhaiXmax.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên TRÁI
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengPhaiXmax);
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương X
                            pointListPartDimY.Add(tamNghiengPhaiXmax);
                            pointListPartDimY.Add(tamNghiengPhaiYmin);
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 200);
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengPhaiXmax);
                                listPointBoltDim_X.Add(tamNghiengPhaiYmin);
                                foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiXmax, tamNghiengPhaiYmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }
                        else if (tamNghiengPhaiXmin.X > mainPartPointMidTop.X && tamNghiengPhaiXmin.Y > mainPartPointMidTop.Y && tamNghiengPhaiXmin.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên PHẢI
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengPhaiXmin);
                            createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 250);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương X
                            pointListPartDimY.Add(tamNghiengPhaiYmax);
                            pointListPartDimY.Add(tamNghiengPhaiXmin);
                            createDimension.CreateStraightDimensionSet_Y(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 200);
                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengPhaiYmax);
                                listPointBoltDim_X.Add(tamNghiengPhaiXmin);
                                foreach (tsm.BoltGroup bgr in tamNghiengPhai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengPhaiYmax, tamNghiengPhaiXmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }

                    }//Duyet tat cả các part cua  partsNghiengPhai

                    foreach (var tamNghiengTrai in partsNghiengTrai)//Duyet tat cả các part cua  partsNghiengTrai
                    {
                        tsd.PointList pointListDimXTamNghieng = new PointList(); // ds điểm để dim phương X cho tam nghieng
                        tsd.PointList pointListDimYTamNghieng = new PointList(); // ds điểm để dim phương Y cho tam nghieng
                        List<t3d.Point> pointListBolt = new List<t3d.Point>(); // ds chưa tất cả điểm của bolts lấy từ part
                        Part_Edge part_EdgeNghiengTrai = new Part_Edge(view, tamNghiengTrai); // ds chưa tất cả điểm part edge cua  tamNghiengPhai
                        List<t3d.Point> listpartEdgeTamThayBoltTron = part_EdgeNghiengTrai.List_Edge;//ds điểm của part 
                        t3d.Point tamNghiengTraiXmin = part_EdgeNghiengTrai.PointXmin0;
                        t3d.Point tamNghiengTraiXmax = part_EdgeNghiengTrai.PointXmax0;
                        t3d.Point tamNghiengTraiYmin = part_EdgeNghiengTrai.PointYmin0;
                        t3d.Point tamNghiengTraiYmax = part_EdgeNghiengTrai.PointYmax0;

                        if (Math.Abs(tamNghiengTraiYmin.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengTrai nằm bên TRÊN
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengTraiYmin);
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                            pointListPartDimX.Add(tamNghiengTraiXmin);
                            pointListPartDimX.Add(tamNghiengTraiYmin);
                            createDimension.CreateStraightDimensionSet_X(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 150);

                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengTraiXmin);
                                listPointBoltDim_X.Add(tamNghiengTraiYmin);
                                foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmin, tamNghiengTraiYmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }

                        }
                        else if (Math.Abs(tamNghiengTraiYmax.Y - mainPartPointMidTop.Y) < 3) //tấm NghiengPhai nằm bên DƯỚI
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengTraiYmax);
                            createDimension.CreateStraightDimensionSet_X_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimX = new PointList(); // ds điểm để dim phương X
                            pointListPartDimX.Add(tamNghiengTraiXmax);
                            pointListPartDimX.Add(tamNghiengTraiYmax);
                            createDimension.CreateStraightDimensionSet_X_(viewBase, pointListPartDimX, cb_DimentionAttributes.Text, 150);

                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                listPointBoltDim_X.Add(tamNghiengTraiYmax);
                                foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiYmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }
                        else if (tamNghiengTraiXmax.X < mainPartPointMidTop.X && tamNghiengTraiXmax.Y > mainPartPointMidTop.Y && tamNghiengTraiXmax.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên TRÁI
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengTraiXmax);
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương Y
                            pointListPartDimY.Add(tamNghiengTraiYmax);
                            pointListPartDimY.Add(tamNghiengTraiXmax);
                            createDimension.CreateStraightDimensionSet_Y_(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 150);

                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                listPointBoltDim_X.Add(tamNghiengTraiYmax);
                                foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiYmax, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }
                        else if (tamNghiengTraiXmin.X > mainPartPointMidTop.X && tamNghiengTraiXmin.Y > mainPartPointMidTop.Y && tamNghiengTraiXmin.Y < mainPartPointMidTop.Y) //tấm NghiengPhai nằm bên PHẢI
                        {
                            pointListDimXTamNghieng.Add(mainPartPointMidTop);
                            pointListDimXTamNghieng.Add(tamNghiengTraiXmin);
                            createDimension.CreateStraightDimensionSet_Y(viewBase, pointListDimXTamNghieng, cb_DimentionAttributes.Text, 350);
                            //Đoạn dưới để xác định góc nghiêng và dim bolt
                            tsd.PointList pointListPartDimY = new PointList(); // ds điểm để dim phương Y
                            pointListPartDimY.Add(tamNghiengTraiXmax);
                            pointListPartDimY.Add(tamNghiengTraiXmin);
                            createDimension.CreateStraightDimensionSet_Y(viewBase, pointListPartDimY, cb_DimentionAttributes.Text, 150);

                            if (chbox_DimInternalPart.Checked == true)
                            {
                                listPointBoltDim_X.Clear();
                                listPointBoltDim_X.Add(tamNghiengTraiXmax);
                                listPointBoltDim_X.Add(tamNghiengTraiXmin);
                                foreach (tsm.BoltGroup bgr in tamNghiengTrai.GetBolts())
                                {
                                    foreach (t3d.Point p in bgr.BoltPositions)
                                    {
                                        listPointBoltDim_X.Add(p);
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(tamNghiengTraiXmax, tamNghiengTraiXmin, viewBase, listPointBoltDim_X, cb_DimentionAttributes.Text, 75);//tạo dim Bolt
                            }
                        }
                    }//Duyet tat cả các part cua  partsNghiengTrai

                }
                catch (Exception)
                { }
            }
            drawingHandler.GetActiveDrawing().CommitChanges();
        }



        private void btn_OpenActiveDrawing_Click_1(object sender, EventArgs e)//button OpenActiveDrawing mở bản vẽ đã ban hành (Export)
        {
            string[] directoryFull = null; //đường dẫn để mở file
            string filePath = string.Empty; //đường dẫn trong ô người dùng nhập
            string modelPath = model.GetInfo().ModelPath;
            string Mark = string.Empty; // Tên bản vẽ Assembly

            string lastRevDr = string.Empty;
            try
            {
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                if (actived_dr == null) goto SelectModelObject;
                // các bước sau đây dể lấy assembly position vì dr.Mark có chứa '[' ']' nên lấy actived_dr.Mark bị sai
                if (actived_dr is tsd.AssemblyDrawing)
                {
                    tsd.AssemblyDrawing assemblyDrawing = actived_dr as tsd.AssemblyDrawing;
                    //Thong qua ban ve ta co the lay assembly ngoai model
                    tsm.Assembly assembly = model.SelectModelObject(assemblyDrawing.AssemblyIdentifier) as tsm.Assembly;
                    //Lay duoc mainPrefix
                    assembly.GetReportProperty("ASSEMBLY_POS", ref Mark);
                    if (txt_RevisionDrawing.Text == "")
                        assembly.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                    else
                        lastRevDr = txt_RevisionDrawing.Text;
                }
                else if (actived_dr is tsd.SinglePartDrawing)
                {
                    tsd.SinglePartDrawing singleDrawing = actived_dr as tsd.SinglePartDrawing;
                    tsm.Part part = model.SelectModelObject(singleDrawing.PartIdentifier) as tsm.Part;
                    part.GetReportProperty("PART_POS", ref Mark);
                    Mark = part.GetPartMark();
                    if (txt_RevisionDrawing.Text == "")
                        part.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                    else
                        lastRevDr = txt_RevisionDrawing.Text;

                }
                else if (actived_dr is tsd.GADrawing)
                {
                    tsd.GADrawing gADrawing = actived_dr as tsd.GADrawing;
                    if (tbx_DirectoryOpenDrawing.Text == "")
                        filePath = modelPath;
                    else
                    {
                        filePath = tbx_DirectoryOpenDrawing.Text;
                        if (filePath.Last() != '\\')
                            filePath = filePath + '\\';
                    }
                    String gaTitle1 = gADrawing.Title1;
                    lastRevDr = txt_RevisionDrawing.Text;
                    if (txt_RevisionDrawing.Text == "")
                    {
                        if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                            directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*.pdf", gaTitle1), SearchOption.AllDirectories);//PDF đường dẫn chứa Title1 và revision
                        else
                            directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*.dwg", gaTitle1), SearchOption.AllDirectories);//CAD đường dẫn chứa Title1 và revision
                                                                                                                                            // MessageBox.Show(directory[0]);
                    }
                    else if (txt_RevisionDrawing.Text != "")
                    {
                        if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                            directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", gaTitle1, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Title1 và revision
                        else
                            directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", gaTitle1, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Title1 và revision
                                                                                                                                                          // MessageBox.Show(directory[0]);
                    }
                    Process.Start(directoryFull[0]);//mở file đầu tiên trong array list directoryFull
                }
                if (tbx_DirectoryOpenDrawing.Text == "")
                    filePath = modelPath;
                else
                {
                    filePath = tbx_DirectoryOpenDrawing.Text;
                    if (filePath.Last() != '\\')
                        filePath = filePath + '\\';
                }
                //    directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", drAsMark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                if (txt_RevisionDrawing.Text == "")
                {
                    if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", Mark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                    else
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", Mark, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Mark và revision
                                                                                                                                                  // MessageBox.Show(directory[0]);
                }
                else if (txt_RevisionDrawing.Text != "")
                {
                    if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", Mark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                    else
                        directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", Mark, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Mark và revision

                }
                Process.Start(directoryFull[0]);//mở file đầu tiên trong array list directoryFull


            SelectModelObject:
                {
                    tsmui.ModelObjectSelector MySelector = new Tekla.Structures.Model.UI.ModelObjectSelector();
                    ModelObjectEnumerator MyObjs = MySelector.GetSelectedObjects();
                    while (MyObjs.MoveNext())
                    {
                        if (MyObjs.Current is tsm.Part)
                        {
                            tsm.Part part = MyObjs.Current as tsm.Part;
                            part.GetReportProperty("PART_POS", ref Mark);
                            if (txt_RevisionDrawing.Text == "")
                                part.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                            else
                                lastRevDr = txt_RevisionDrawing.Text;
                        }
                        else
                        {
                            tsm.Assembly asbl = MyObjs.Current as tsm.Assembly;
                            asbl.GetReportProperty("ASSEMBLY_POS", ref Mark);
                            if (txt_RevisionDrawing.Text == "")
                                asbl.GetReportProperty("DRAWING.REVISION.MARK", ref lastRevDr);
                            else
                                lastRevDr = txt_RevisionDrawing.Text;
                        }

                        if (tbx_DirectoryOpenDrawing.Text == "")
                            filePath = modelPath;
                        else
                        {
                            filePath = tbx_DirectoryOpenDrawing.Text;
                            if (filePath.Last() != '\\')
                                filePath = filePath + '\\';
                        }
                        if (txt_RevisionDrawing.Text == "")
                        {
                            if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                                directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", Mark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                            else
                                directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", Mark, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Mark và revision
                                                                                                                                                          // MessageBox.Show(directory[0]);
                        }
                        else if (txt_RevisionDrawing.Text != "")
                        {
                            if (cb_PdfFileOpen.Checked) //Nếu check vào ô mở PDF
                                directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.pdf", Mark, lastRevDr), SearchOption.AllDirectories);//PDF đường dẫn chứa Mark và revision
                            else
                                directoryFull = Directory.GetFiles(filePath, string.Format("*{0}*{1}.dwg", Mark, lastRevDr), SearchOption.AllDirectories);//CAD đường dẫn chứa Mark và revision

                        }
                        Process.Start(directoryFull[0]);//mở file đầu tiên trong array list directoryFull  directoryFull.Length - 1
                    }
                }

            }
            catch
            {
                MessageBox.Show("Bản vẽ không tìm thấy, thử nhập số rev bản vẽ vào ô Rev rồi mở lại.");
            }
        }

        private void btn_Change_Click_1(object sender, EventArgs e)
        {
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            string revText = txt_RevText.Text;
            string datenow = DateTime.Now.ToString("MM.dd");
            if (revText == "")
                revText = "1R" + datenow;
            if (cb_SelectTitleReviseText.SelectedIndex == -1 || cb_SelectTitleReviseText.SelectedIndex == 2)
                actived_dr.Title3 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 0)
                actived_dr.Title1 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 1)
                actived_dr.Title2 = revText;
            actived_dr.Modify();
            drawingHandler.SaveActiveDrawing();
        }

        private void btn_Decrease_Click(object sender, EventArgs e)
        {
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            string revText = txt_RevText.Text;
            string datenow = DateTime.Now.ToString("MM.dd");
            if (revText == "")
                revText = "1R" + datenow + "-D";

            if (cb_SelectTitleReviseText.SelectedIndex == -1 || cb_SelectTitleReviseText.SelectedIndex == 2)
                actived_dr.Title3 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 0)
                actived_dr.Title1 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 1)
                actived_dr.Title2 = revText;
            actived_dr.Modify();
            drawingHandler.SaveActiveDrawing();
        }

        private void btn_Increase_Click(object sender, EventArgs e)
        {
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            string revText = txt_RevText.Text;
            string datenow = DateTime.Now.ToString("MM.dd");
            if (revText == "")
                revText = "1R" + datenow + "-I";

            if (cb_SelectTitleReviseText.SelectedIndex == -1 || cb_SelectTitleReviseText.SelectedIndex == 2)
                actived_dr.Title3 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 0)
                actived_dr.Title1 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 1)
                actived_dr.Title2 = revText;
            actived_dr.Modify();
            drawingHandler.SaveActiveDrawing();
        }

        private void btn_New_Click(object sender, EventArgs e)
        {
            tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
            string revText = txt_RevText.Text;
            string datenow = DateTime.Now.ToString("MM.dd");
            if (revText == "")
                revText = "1R" + datenow + "-N";

            if (cb_SelectTitleReviseText.SelectedIndex == -1 || cb_SelectTitleReviseText.SelectedIndex == 2)
                actived_dr.Title3 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 0)
                actived_dr.Title1 = revText;
            else if (cb_SelectTitleReviseText.SelectedIndex == 1)
                actived_dr.Title2 = revText;
            actived_dr.Modify();
            drawingHandler.SaveActiveDrawing();
        }

        private void btn_ClearTitle3_Click(object sender, EventArgs e)
        {
            if (Control.ModifierKeys == Keys.Shift)
            {
                try
                {
                    tsd.DrawingObjectEnumerator drObjEnum = drawingHandler.GetActiveDrawing().GetSheet().GetAllObjects();
                    while (drObjEnum.MoveNext())
                    {
                        if (drObjEnum.Current is tsd.Plugin)
                        {
                            tsd.Plugin plugin = drObjEnum.Current as tsd.Plugin;
                            if (plugin.Name == "CNCloudPlugin")
                            {
                                plugin.Delete();
                            }
                        }
                        else if (drObjEnum.Current is tsd.Plugin)
                        {
                            tsd.View view = drObjEnum.Current as tsd.View;
                            tsd.DrawingObjectEnumerator pluginEnum = view.GetAllObjects(new Type[] { typeof(tsd.Plugin) });
                            foreach (tsd.Plugin plugin in pluginEnum)
                            {
                                if (plugin.Name == "CNCloudPlugin")
                                {
                                    plugin.Delete();
                                }
                            }
                        }
                    }
                    drawingHandler.GetActiveDrawing().CommitChanges();
                }
                catch
                { Operation.DisplayPrompt("Not found CNCloudPlugin"); }
            }
            else
            {
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                actived_dr.Title3 = "";
                actived_dr.Modify();
                drawingHandler.SaveActiveDrawing();
            }
        }



        private void button1_Click(object sender, EventArgs e) //nút khởi động lại form
        {
            try
            {
                Application.Restart();
                Environment.Exit(0);
                frm_Main.ActiveForm.Height = 460;
                frm_Main.ActiveForm.Width = 285;
                //Lay duong dan cua model dang duoc mo
                string modelPath = model.GetInfo().ModelPath;
                string attributePath = Path.Combine(modelPath, "attributes");
                //Lay cac file o trong attribute folder
                string[] drAsAttributes = Directory.GetFiles(attributePath, "*.ad"); //load attribute bản vẽ assembly trong folder Model
                string[] drSgAttributes = Directory.GetFiles(attributePath, "*.wd"); //load attribute bản vẽ assembly trong folder Model
                string[] drDimAttributes = Directory.GetFiles(attributePath, "*.dim"); //load attribute DIM trong folder Model


                if (modelPath.Last() != '\\')
                    modelPath = modelPath + '\\';
                tbx_DirectoryOpenDrawing.Text = modelPath; //Hiển thị đường dẫn model ở ô textbox tab Numbering drawing, để mở drawing active
                                                           //Tạo ra 1 list phụ để kiểm tra phần tử trùng
                List<string> singlePartAttributes = new List<string>();
                List<string> assemblyAttributes = new List<string>();
                foreach (var item in drDimAttributes)
                {
                    cb_DimentionAttributes.Items.Add(Path.GetFileNameWithoutExtension(item)); //đưa dữ liệu vào ô DimentionProperties.
                    cb_DimAttDimPartBoltAssembly.Items.Add(Path.GetFileNameWithoutExtension(item)); //đưa dữ liệu vào ô cb_DimAttDimPartBoltAssembly của button dim parts/bolts of assembly
                }
            }
            //((DataGridViewComboBoxColumn)(dtgv_AssDrawingAttribute.Columns[dt_AssDrAttribute.Index])).DataSource = cb_ASDrAttributes.Items;

            catch { }
        }

        private void btn_DimPartOverallBolt2_Click(object sender, EventArgs e)
        {
            try
            {
                if (cb_PartBoltOption.SelectedIndex == 1)
                    DimGratingCheckerPlate();//Nếu check vào ô dim grating thì gọi hàm DimGratingCheckerPlate
                                             // Khai báo drawingHandler và model đã làm rồi (pulic static), nếu chưa có có thể thêm ở đây.
                                             //Khai báo workplanhandler
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsdui.Picker c_picker = drawingHandler.GetPicker();
                int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drpart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mpart = model.SelectModelObject(drpart.ModelIdentifier) as tsm.Part;
                        //Lấy chiều dày và profile của tấm
                        double cday = 0;
                        string partProfileType = null;
                        mpart.GetReportProperty("PROFILE.WIDTH", ref cday);
                        mpart.GetReportProperty("PROFILE_TYPE", ref partProfileType);

                        //Lấy view hiện hành thông qua drpart
                        tsd.View viewcur = drpart.GetView() as tsd.View;
                        tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        tsd.ViewBase viewbase = drpart.GetView();

                        //Tính điểm định nghĩa của part theo bản vẽ
                        Part_Edge part_edge = new Part_Edge(viewcur, mpart);

                        t3d.Point minx = part_edge.PointXmin0;
                        t3d.Point maxx = part_edge.PointXmax0;
                        t3d.Point miny = part_edge.PointYmin0;
                        t3d.Point maxy = part_edge.PointYmax0;

                        //MessageBox.Show(string.Format("minX {0}{1} maxX {2}{3} minY {4}{5} maxY {6}{7}", minX_.ToString(), minx.ToString(), maxX_.ToString(), maxx.ToString(), minY_.ToString(), miny.ToString(), maxY_.ToString(), maxy.ToString()));

                        if (Math.Abs(maxx.X - minx.X - cday) < 3 && partProfileType == "B")// Tấm đứng (front view), profile Plate
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();//Ds 2 điểm thấp nhất và cao nhất bên trái
                            pl.Add(new t3d.Point(part_edge.PointXminYmin));//add thêm Điểm có tọa độ X nhỏ nhất và tọa độ Y nhỏ nhất
                            pl.Add(new t3d.Point(part_edge.PointXminYmax));//add thêm Điểm có tọa độ X nhỏ nhất và tọa độ Y lớn nhất
                            bolt_pl.Add(new t3d.Point(part_edge.PointXminYmin));
                            bolt_pl.Add(new t3d.Point(part_edge.PointXminYmax));
                            if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                            {
                                createDimension.CreateStraightDimensionSet_Y_(viewbase, pl, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương -Y
                            }
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương -Y
                        } // Tấm đứng (front view), profile Plate
                        else if (Math.Abs(maxy.Y - miny.Y - cday) < 3 && partProfileType == "B")// Tấm ngang (front vỉew)
                        {
                            //MessageBox.Show(tietdien.ToString());
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();//Ds 2 điểm thấp nhất bên dưới
                            pl.Add(new t3d.Point(part_edge.PointXminYmin)); //add thêm Điểm có tọa độ X nhỏ nhất và tọa độ Y nhỏ nhất
                            pl.Add(new t3d.Point(part_edge.PointXmaxYmin)); //add thêm Điểm có tọa độ X lớn nhất và tọa độ Y nhỏ nhất
                            bolt_pl.Add(new t3d.Point(part_edge.PointXminYmin));
                            bolt_pl.Add(new t3d.Point(part_edge.PointXmaxYmin));
                            if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                            {
                                createDimension.CreateStraightDimensionSet_X_(viewbase, pl, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương -X
                            }
                            createDimension.CreateStraightDimensionSet_X_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương -X
                        } // Tấm ngang (front vỉew)
                        else if (Math.Abs(t3d.Distance.PointToPoint(minx, miny) - cday) < 3 && partProfileType == "B")// Tấm nghiêng phải
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();
                            pl.Add(miny);//Điểm có tọa độ Y nhỏ nhất
                            pl.Add(maxx);//Điểm có tọa độ X lớn nhất
                            bolt_pl.Add(miny);
                            bolt_pl.Add(maxx);
                            t3d.Vector vx = new Vector(maxx.X - miny.X, maxx.Y - miny.Y, 0);
                            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                            if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                            {
                                createDimension.CreateStraightDimensionSet_FX_(miny, maxx, viewbase, pl, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương F
                            }
                            createDimension.CreateStraightDimensionSet_FX_(miny, maxx, viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương F
                        } // Tấm nghiêng phải
                        else if (Math.Abs(t3d.Distance.PointToPoint(miny, maxx) - cday) < 3 && partProfileType == "B")// Tấm nghiêng trái
                        {
                            tsd.PointList bolt_pl = new PointList();
                            foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                            {
                                foreach (t3d.Point p in bgr.BoltPositions)
                                {
                                    bolt_pl.Add(p);
                                }
                            }
                            tsd.PointList pl = new PointList();
                            pl.Add(miny);//Điểm có tọa độ Y nhỏ nhất
                            pl.Add(minx);//Điểm có tọa độ X nhỏ nhất
                            bolt_pl.Add(miny);
                            bolt_pl.Add(minx);
                            t3d.Vector vx = new Vector(miny.X - minx.X, miny.Y - minx.Y, 0);
                            t3d.Vector vy = vx.Cross(new t3d.Vector(0, 0, 1));
                            if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                            {
                                createDimension.CreateStraightDimensionSet_FX(miny, minx, viewbase, pl, cb_DimentionAttributes.Text, dimSpace);
                            } //Tạo đường dim tổng
                            createDimension.CreateStraightDimensionSet_FX(miny, minx, viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương . Dùng CreateStraightDimensionSet_FX_ nếu muốn đặt dim ở hướng ngược lại
                        } // Tấm nghiêng trái

                        else if (partProfileType == "I" || partProfileType == "L" || partProfileType == "U" || partProfileType == "T" || partProfileType == "C" || partProfileType == "Z" || partProfileType == "RO" || partProfileType == "RU" || partProfileType == "M")
                        {
                            tsd.PointList bolt_pl = new PointList();
                            ArrayList arrayListCenterLine = mpart.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            //MessageBox.Show(pointCenterLineXmin.ToString() + "   " + pointCenterLineXmax.ToString());
                            t3d.Point pointXmax = part_edge.PointXmax0;
                            t3d.Point pointXmin = part_edge.PointXmin0;
                            t3d.Point pointYmax = part_edge.PointYmax0;
                            t3d.Point pointYmin = part_edge.PointYmin0;
                            t3d.Point pointXminYmin = part_edge.PointXminYmin;
                            t3d.Point pointXminYmax = part_edge.PointXminYmax;
                            t3d.Point pointXmaxYmax = part_edge.PointXmaxYmax;
                            t3d.Point pointXmaxYmin = part_edge.PointXmaxYmin;
                            tsd.PointList listpointPartX = new PointList();
                            tsd.PointList listpointPartY_ = new PointList();
                            tsd.PointList listpointBoltX = new PointList();
                            tsd.PointList listpointBoltY_ = new PointList();
                            tsd.PointList listpointBoltY = new PointList();

                            if (pointCenterLineXmin.Y < pointCenterLineXmax.Y && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 2
                                && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part nằm nghiêng Phải
                            {
                                listpointPartX.Add(pointCenterLineXmin);
                                listpointPartX.Add(pointCenterLineXmax);
                                listpointPartY_.Add(pointYmax);
                                listpointPartY_.Add(pointXmax);
                                listpointBoltX.Add(pointCenterLineXmin);
                                listpointBoltX.Add(pointCenterLineXmax);
                                listpointBoltY_.Add(pointYmax);
                                listpointBoltY_.Add(pointXmax);

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointPartX, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương FX_
                                    createDimension.CreateStraightDimensionSet_FY_(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointPartY_, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương FY
                                }

                                foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY;
                                    if (pointListBolt_XY != null)
                                    {
                                        foreach (t3d.Point p in pointListBolt_XY)
                                        {
                                            listpointBoltX.Add(p);
                                            listpointBoltY_.Add(p);
                                        }
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointBoltX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FX
                                createDimension.CreateStraightDimensionSet_FY_(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointBoltY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FY
                            } //Part nằm nghiêng Phải
                            else if (pointCenterLineXmin.Y > pointCenterLineXmax.Y && Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) > 2
                                && Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) > 2) //Part nằm nghiêng Trái
                            {
                                listpointPartX.Add(pointCenterLineXmin);
                                listpointPartX.Add(pointCenterLineXmax);
                                listpointPartY_.Add(pointYmax);
                                listpointPartY_.Add(pointXmin);
                                listpointBoltX.Add(pointCenterLineXmin);
                                listpointBoltX.Add(pointCenterLineXmax);
                                listpointBoltY_.Add(pointYmax);
                                listpointBoltY_.Add(pointXmin);

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointPartX, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương FX
                                    createDimension.CreateStraightDimensionSet_FY(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointPartY_, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương FY
                                }
                                foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY;
                                    if (pointListBolt_XY != null)
                                    {
                                        foreach (t3d.Point p in pointListBolt_XY)
                                        {
                                            listpointBoltX.Add(p);
                                            listpointBoltY_.Add(p);
                                        }
                                    }
                                }
                                createDimension.CreateStraightDimensionSet_FX(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointBoltX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FX
                                createDimension.CreateStraightDimensionSet_FY(pointCenterLineXmin, pointCenterLineXmax, viewbase, listpointBoltY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương FY
                            } //Part nằm nghiêng Trái

                            else if (Math.Abs(pointCenterLineXmax.Y - pointCenterLineXmin.Y) < 2) //Part nằm NGANG
                            {
                                double distanceBolt_PartXmin = 0;
                                double distanceBolt_PartXmax = 0;
                                PointList listpointBoltXYdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBoltXYdimY_ = new PointList() { pointXminYmax, pointXminYmin };
                                PointList listpointBoltXYdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };

                                PointList listpointBolt_YdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBolt_YdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };

                                PointList listpointBolt_XdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };
                                PointList listpointBolt_XdimY_ = new PointList() { pointXminYmax, pointXminYmin };

                                listpointPartX.Add(pointXminYmin);
                                listpointPartX.Add(pointXmaxYmin);
                                listpointPartY_.Add(pointXminYmax);
                                listpointPartY_.Add(pointXminYmin);

                                listpointBoltX.Add(pointXminYmin);
                                listpointBoltX.Add(pointXmaxYmin);
                                listpointBoltY_.Add(pointXminYmax);
                                listpointBoltY_.Add(pointXminYmin);
                                listpointBoltY.Add(pointXmaxYmax);
                                listpointBoltY.Add(pointXmaxYmin);
                                tsm.BoltGroup boltGroupXY = null;
                                tsm.BoltGroup boltGroupY = null;
                                tsm.BoltGroup boltGroupX = null;

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewbase, listpointPartX, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointPartY_, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương Y
                                }
                                List<BoltGroup> listBoltGroupXY = new List<BoltGroup>(); // list chứa các boltgroup Tròn
                                List<BoltGroup> listBoltGroupY = new List<BoltGroup>();// list chứa các boltgroup Nằm Dọc
                                List<BoltGroup> listBoltGroupX = new List<BoltGroup>();// list chứa các boltgroup Nằm Ngang

                                foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    boltGroupXY = tinh_Toan_Bolt.Bolt_XY; // 1 boltgroup Tròn
                                    boltGroupY = tinh_Toan_Bolt.Bolt_Y;// 1 boltgroup Nằm Dọc
                                    boltGroupX = tinh_Toan_Bolt.Bolt_X; // 1 boltgroup Nằm Ngang

                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY; // ds tất cả các điểm của boltgroup Tròn
                                    List<t3d.Point> pointListBolt_Y = tinh_Toan_Bolt.PointListBolt_Y; // ds tất cả các điểm của boltgroup Nằm Dọc
                                    List<t3d.Point> pointListBolt_X = tinh_Toan_Bolt.PointListBolt_X; // ds tất cả các điểm của boltgroup Ngang

                                    if (boltGroupXY != null) //nếu bolt Tròn không rỗng, xét TH có 1, 2, n bolt group
                                    {
                                        listBoltGroupXY.Add(boltGroupXY);
                                        foreach (t3d.Point p in pointListBolt_XY)
                                        {
                                            distanceBolt_PartXmin = Distance.PointToPoint(p, pointXminYmax); // thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim Y qua trái và phải là xong.
                                            distanceBolt_PartXmax = Distance.PointToPoint(p, pointXmaxYmax);// thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim Y qua trái và phải là xong.
                                            listpointBoltXYdimX.Add(p);
                                            listpointBoltXYdimY_.Add(p);
                                            listpointBoltXYdimY.Add(p);
                                        }
                                    }

                                    else if (boltGroupY != null)
                                    {
                                        listBoltGroupY.Add(boltGroupY);
                                        foreach (t3d.Point p in pointListBolt_Y)
                                        {
                                            if (p.Y > pointCenterLineXmin.Y) //Nếu nằm TRÊN đường centerline
                                            {
                                                listpointBolt_YdimX.Add(p);
                                            }
                                            else //Nếu nằm DƯỚI đường centerline
                                            {
                                                listpointBolt_YdimX_.Add(p);
                                            }
                                        }
                                    }

                                    else if (pointListBolt_X != null)
                                    {
                                        MessageBox.Show("Có bolt nằm ngang");
                                        break;
                                    }
                                }

                                //DIM cho bolt TRÒN
                                createDimension.CreateStraightDimensionSet_X(viewbase, listpointBoltXYdimX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt TRÒN
                                if (listBoltGroupXY.Count == 1) // nếu chỉ có 1 BoltGroup TRÒN
                                {
                                    if (distanceBolt_PartXmin < distanceBolt_PartXmax) //Nếu kc từ bolt đến mép trái part nhỏ hơn, thì tạo dim qua trái
                                        createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointBoltXYdimY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua trái
                                    else
                                        createDimension.CreateStraightDimensionSet_Y(viewbase, listpointBoltXYdimY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua phải
                                }
                                else if (listBoltGroupXY.Count == 2) // nếu có 2 BoltGroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointBoltXYdimY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua trái
                                    createDimension.CreateStraightDimensionSet_Y(viewbase, listpointBoltXYdimY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua phải
                                }
                                else //nếu có nhiều boltgroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointBoltXYdimY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim);// tạo DIM Bolt phương Y qua trái
                                }
                                //DIM cho bolt NẰM DỌC
                                if (boltGroupY != null)
                                {
                                    if (listpointBolt_YdimX.Count > 2)
                                        createDimension.CreateStraightDimensionSet_X(viewbase, listpointBolt_YdimX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim - 25); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                    if (listpointBolt_YdimX_.Count > 2)
                                        createDimension.CreateStraightDimensionSet_X_(viewbase, listpointBolt_YdimX_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt NẰM ODC
                                }

                            }
                            else if (Math.Abs(pointCenterLineXmax.X - pointCenterLineXmin.X) < 2) //Part ĐỨNG
                            {
                                double distanceBolt_PartYmin = 0;
                                double distanceBolt_PartYmax = 0;
                                PointList listpointBoltXYdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBoltXYdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };
                                PointList listpointBoltXYdimY_ = new PointList() { pointXminYmax, pointXminYmin };
                                PointList listpointBoltXYdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };

                                PointList listpointBolt_YdimX = new PointList() { pointXminYmax, pointXmaxYmax };
                                PointList listpointBolt_YdimX_ = new PointList() { pointXminYmin, pointXmaxYmin };

                                PointList listpointBolt_XdimY = new PointList() { pointXmaxYmax, pointXmaxYmin };
                                PointList listpointBolt_XdimY_ = new PointList() { pointXminYmax, pointXminYmin };

                                listpointPartX.Add(pointXminYmax);
                                listpointPartX.Add(pointXmaxYmax);
                                listpointPartY_.Add(pointXminYmax);
                                listpointPartY_.Add(pointXminYmin);

                                listpointBoltX.Add(pointXminYmax);
                                listpointBoltX.Add(pointXmaxYmax);
                                listpointBoltY_.Add(pointXminYmax);
                                listpointBoltY_.Add(pointXminYmin);
                                listpointBoltY.Add(pointXmaxYmax);
                                listpointBoltY.Add(pointXmaxYmin);
                                tsm.BoltGroup boltGroupXY = null;
                                tsm.BoltGroup boltGroupY = null;
                                tsm.BoltGroup boltGroupX = null;
                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewbase, listpointPartX, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointPartY_, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương Y
                                }
                                List<BoltGroup> listBoltGroupXY = new List<BoltGroup>(); // list chứa các boltgroup Tròn
                                List<BoltGroup> listBoltGroupY = new List<BoltGroup>();// list chứa các boltgroup Nằm Dọc
                                List<BoltGroup> listBoltGroupX = new List<BoltGroup>();// list chứa các boltgroup Nằm Ngang

                                foreach (tsm.BoltGroup bgr in mpart.GetBolts())
                                {
                                    Tinh_Toan_Bolt tinh_Toan_Bolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                    boltGroupXY = tinh_Toan_Bolt.Bolt_XY; // 1 boltgroup Tròn
                                    boltGroupY = tinh_Toan_Bolt.Bolt_Y;// 1 boltgroup Nằm Dọc
                                    boltGroupX = tinh_Toan_Bolt.Bolt_X; // 1 boltgroup Nằm Ngang

                                    List<t3d.Point> pointListBolt_XY = tinh_Toan_Bolt.PointListBolt_XY; // ds tất cả các điểm của boltgroup Tròn
                                    List<t3d.Point> pointListBolt_Y = tinh_Toan_Bolt.PointListBolt_Y; // ds tất cả các điểm của boltgroup Nằm Dọc
                                    List<t3d.Point> pointListBolt_X = tinh_Toan_Bolt.PointListBolt_X; // ds tất cả các điểm của boltgroup Ngang

                                    if (boltGroupXY != null) //nếu bolt Tròn không rỗng, xét TH có 1, 2, n bolt group
                                    {
                                        listBoltGroupXY.Add(boltGroupXY);
                                        foreach (t3d.Point p in pointListBolt_XY)
                                        {
                                            distanceBolt_PartYmin = Distance.PointToPoint(p, pointXminYmin); // thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim X lên trên xuống dưới.
                                            distanceBolt_PartYmax = Distance.PointToPoint(p, pointXminYmax);// thông số này để xét đến khi chỉ có 1 BoltGroup XY, TH có 2 thì chỉ cần tạo ra 2 dim X lên trên xuống dưới.
                                            listpointBoltXYdimX.Add(p);
                                            listpointBoltXYdimX_.Add(p);
                                            listpointBoltXYdimY_.Add(p);
                                        }
                                    }

                                    else if (boltGroupX != null)
                                    {
                                        listBoltGroupY.Add(boltGroupY);
                                        foreach (t3d.Point p in pointListBolt_X)
                                        {
                                            if (p.X < pointCenterLineXmin.X) //Nếu nằm BÊN TRÁI đường centerline
                                            {
                                                listpointBolt_XdimY_.Add(p);
                                            }
                                            else //Nếu nằm BÊN PHẢI đường centerline
                                            {
                                                listpointBolt_XdimY.Add(p);
                                            }

                                        }
                                    }

                                    else if (boltGroupY != null)
                                    {
                                        MessageBox.Show("Có bolt nằm ngang");
                                    }
                                }

                                //DIM cho bolt TRÒN
                                createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointBoltXYdimY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X cho bolt TRÒN
                                if (listBoltGroupXY.Count == 1) // nếu chỉ có 1 BoltGroup TRÒN
                                {
                                    if (distanceBolt_PartYmin > distanceBolt_PartYmax) //Nếu kc từ bolt đến mép DƯỚI part lớn hơn, thì tạo dim lên TRÊN
                                        createDimension.CreateStraightDimensionSet_X(viewbase, listpointBoltXYdimX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim);   // tạo DIM Bolt phương X lên trên
                                    else
                                        createDimension.CreateStraightDimensionSet_X_(viewbase, listpointBoltXYdimX_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X xuống dưới
                                }
                                else if (listBoltGroupXY.Count == 2) // nếu có 2 BoltGroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewbase, listpointBoltXYdimX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); //   tạo DIM Bolt phương X qua lên trên
                                    createDimension.CreateStraightDimensionSet_X_(viewbase, listpointBoltXYdimX_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương X qua xuống dưới

                                }
                                else //nếu có nhiều boltgroup TRÒN
                                {
                                    createDimension.CreateStraightDimensionSet_X(viewbase, listpointBoltXYdimX, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim);// tạo DIM Bolt phương Y qua trái
                                }
                                //DIM cho bolt NẰM NGANG
                                if (boltGroupX != null)
                                {
                                    if (listpointBolt_XdimY_.Count > 2)
                                        createDimension.CreateStraightDimensionSet_Y_(viewbase, listpointBolt_XdimY_, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim - 30); // tạo DIM Bolt phương Y qua trái cho bolt NẰM NGANG
                                    if (listpointBolt_XdimY.Count > 2)
                                        createDimension.CreateStraightDimensionSet_Y(viewbase, listpointBolt_XdimY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM Bolt phương Y qua phải cho bolt NẰM NGANG
                                }


                            } //Part THÉP HÌNH

                            else //Tấm thấy bolt tròn
                            {
                                tsd.PointList dimXPart = new PointList();
                                tsd.PointList dimYPart = new PointList();
                                tsd.PointList Boltlist_Bolt_XY = new PointList();
                                tsd.PointList Boltlist_Bolt_Y = new PointList();
                                tsd.PointList Boltlist_Bolt_X = new PointList();
                                tsd.PointList Boltlist_Bolt_skew = new PointList();
                                //kiểm tra xem đối tượng này có phải là part không ?
                                if (drobj is tsd.Part)
                                {
                                    tsm.ModelObjectEnumerator boltenum = mpart.GetBolts();
                                    foreach (tsm.BoltGroup bgr in boltenum)
                                    {
                                        Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                        //Tìm bolt theo phương XY
                                        tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY;
                                        tsm.BoltGroup bolt_y = tinhbolt.Bolt_Y;
                                        wph.SetCurrentTransformationPlane(viewplane);
                                        if (bolt_xy != null)
                                        {
                                            bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_xy.BoltPositions)
                                            {
                                                Boltlist_Bolt_skew.Add(p);
                                                Boltlist_Bolt_XY.Add(p);
                                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, maxy.Y));
                                                Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, maxy.Y));
                                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));
                                            }
                                        }
                                        if (bolt_y != null)
                                        {
                                            bolt_y.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                            foreach (t3d.Point p in bolt_y.BoltPositions)
                                            {
                                                Boltlist_Bolt_Y.Add(p);
                                                Boltlist_Bolt_Y.Add(new t3d.Point(minx.X, maxy.Y));
                                                Boltlist_Bolt_Y.Add(new t3d.Point(maxx.X, maxy.Y));
                                            }
                                        }
                                    }

                                    dimXPart.Add(new t3d.Point(minx.X, maxy.Y));// góc trên bên trái
                                    dimXPart.Add(new t3d.Point(maxx.X, maxy.Y));// góc trên bên phải
                                    dimYPart.Add(new t3d.Point(minx.X, maxy.Y)); // góc trên bên trái
                                    dimYPart.Add(new t3d.Point(minx.X, miny.Y)); // góc dưới bên trái

                                    if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                    {
                                        createDimension.CreateStraightDimensionSet_X_(viewbase, dimXPart, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương X
                                        createDimension.CreateStraightDimensionSet_Y_(viewbase, dimYPart, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương Y
                                    }
                                    createDimension.CreateStraightDimensionSet_X(viewbase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM BOlt phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương -Y
                                }
                            } //Tấm thấy bolt tròn
                            wph.SetCurrentTransformationPlane(new TransformationPlane());
                        }

                        else if (Math.Abs(maxx.X - minx.X - cday) > 3 && partProfileType == "B") // Tấm THẤY BOLT TRÒN, profile Plate
                        {
                            tsd.PointList bolt_pl = new PointList();
                            ArrayList arrayListCenterLine = mpart.GetCenterLine(true);
                            List<t3d.Point> listCenterLine = new List<t3d.Point> { arrayListCenterLine[0] as t3d.Point, arrayListCenterLine[1] as t3d.Point };
                            t3d.Point pointCenterLineXmin = listCenterLine.OrderBy(point => point.X).ToList()[0];
                            t3d.Point pointCenterLineXmax = listCenterLine.OrderBy(point => point.X).ToList()[1];
                            //MessageBox.Show(pointCenterLineXmin.ToString() + "   " + pointCenterLineXmmax.ToString());
                            t3d.Point pointXmax = part_edge.PointXmax0;
                            t3d.Point pointXmin = part_edge.PointXmin0;
                            t3d.Point pointYmax = part_edge.PointYmax0;
                            t3d.Point pointYmin = part_edge.PointYmin0;
                            t3d.Point pointXminYmin = part_edge.PointXminYmin;
                            t3d.Point pointXminYmax = part_edge.PointXminYmax;
                            t3d.Point pointXmaxYmax = part_edge.PointXmaxYmax;
                            t3d.Point pointXmaxYmin = part_edge.PointXmaxYmin;
                            tsd.PointList listpointPartX = new PointList();
                            tsd.PointList listpointPartY_ = new PointList();
                            tsd.PointList listpointBoltX = new PointList();
                            tsd.PointList listpointBoltY_ = new PointList();
                            tsd.PointList listpointBoltY = new PointList();

                            tsd.PointList dimXPart = new PointList();
                            tsd.PointList dimYPart = new PointList();
                            tsd.PointList Boltlist_Bolt_XY = new PointList();
                            tsd.PointList Boltlist_Bolt_Y = new PointList();
                            tsd.PointList Boltlist_Bolt_X = new PointList();
                            tsd.PointList Boltlist_Bolt_skew = new PointList();
                            tsm.ModelObjectEnumerator boltenum = mpart.GetBolts();
                            foreach (tsm.BoltGroup bgr in boltenum)
                            {
                                dimXPart.Add(new t3d.Point(minx.X, miny.Y));// góc duoi bên trái
                                dimXPart.Add(new t3d.Point(maxx.X, miny.Y));// góc duoi bên phải
                                dimYPart.Add(new t3d.Point(minx.X, maxy.Y)); // góc trên bên trái
                                dimYPart.Add(new t3d.Point(minx.X, miny.Y)); // góc dưới bên trái
                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));// góc trên bên trái
                                Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, miny.Y));// góc trên bên phải
                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));// góc dưới bên trái
                                Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, maxy.Y));// góc dưới bên trái

                                Tinh_Toan_Bolt tinhbolt = new Tinh_Toan_Bolt(viewcur, bgr);
                                //Tìm bolt theo phương XY
                                tsm.BoltGroup bolt_xy = tinhbolt.Bolt_XY;
                                wph.SetCurrentTransformationPlane(viewplane);
                                if (bolt_xy != null)
                                {
                                    bolt_xy.Select();//Chú ý khi đổi hệ tọa độ mà muốn lấy tọa độ của đối tượng cần phải select.
                                    foreach (t3d.Point p in bolt_xy.BoltPositions)
                                    {
                                        Boltlist_Bolt_XY.Add(p);
                                        //Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, maxy.Y));
                                        //Boltlist_Bolt_XY.Add(new t3d.Point(maxx.X, maxy.Y));
                                        //Boltlist_Bolt_XY.Add(new t3d.Point(minx.X, miny.Y));
                                    }
                                }

                                if (cb_PartBoltOption.SelectedIndex == -1 || cb_PartBoltOption.SelectedIndex == 0) //Nếu KHÔNG check vào ô Bolt Only
                                {
                                    createDimension.CreateStraightDimensionSet_X_(viewbase, dimXPart, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương X
                                    createDimension.CreateStraightDimensionSet_Y_(viewbase, dimYPart, cb_DimentionAttributes.Text, dimSpace); // tạo DIM part phương Y
                                }
                                createDimension.CreateStraightDimensionSet_X_(viewbase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM BOlt phương X
                                createDimension.CreateStraightDimensionSet_Y_(viewbase, Boltlist_Bolt_XY, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương -Y
                            }


                        }//Tấm thấy bolt tròn
                        wph.SetCurrentTransformationPlane(new TransformationPlane());
                    }
                    if (drobj is tsd.Bolt)
                    {
                        tsd.Bolt drBolt = drobj as tsd.Bolt;
                        tsd.ViewBase viewbase = drobj.GetView();
                        tsd.View viewcur = drobj.GetView() as tsd.View;
                        wph.SetCurrentTransformationPlane(new TransformationPlane());
                        tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
                        wph.SetCurrentTransformationPlane(viewplane);
                        model.CommitChanges();
                        tsd.PointList bolt_pl = new PointList();
                        //tsm.ModelObject myPart = new tsm.Model().SelectModelObject(modelObject.ModelIdentifier) as tsm.ModelObject;
                        tsm.BoltGroup boltGroup = new tsm.Model().SelectModelObject(drBolt.ModelIdentifier) as tsm.BoltGroup;

                        foreach (t3d.Point p in boltGroup.BoltPositions)
                        {
                            //MessageBox.Show(p.ToString());
                            bolt_pl.Add(p);
                        }

                        try
                        {
                            createDimension.CreateStraightDimensionSet_Y_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương Y
                            createDimension.CreateStraightDimensionSet_X_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); // tạo DIM bolt phương X
                        }
                        catch
                        { createDimension.CreateStraightDimensionSet_X_(viewbase, bolt_pl, cb_DimentionAttributes.Text, dimSpace - dimSpaceDim2Dim); }
                    }
                }
            }
            catch { }
        }
        //public int minFormHeight = 180;
        //public int maxFormHeight = 520;
        //public int midFormHeight = 509;
        //public int minFormWidth = 290;
        //public int maxFormWidth = 577;
        bool gb9 = false;
        private void btn_MoreOption_Click(object sender, EventArgs e)//Nút thay đổi độ cao form
        {
            if (Math.Abs(this.Height - maxFormHeight) <= 5)
            {
                this.Height = minFormHeight; //285
            }
            else if (Math.Abs(this.Height - minFormHeight) <= 5)
            {
                this.Height = midFormHeight;
            }
            else if (Math.Abs(this.Height - midFormHeight) <= 5)
            {
                this.Height = maxFormHeight;
            }
        }

        bool gb8 = false;
        private void button1_Click_1(object sender, EventArgs e) // Nút thay đổi độ rộng form
        {
            if (!gb8)
            {
                frm_Main.ActiveForm.Width = maxFormWidth; //285
            }
            else
            {
                frm_Main.ActiveForm.Width = minFormWidth;
            }
            gb8 = !gb8;
        }




        private void btn_CopySectionDetail_Click(object sender, EventArgs e)
        {

            string error_text = "[ERROR] copying";
            copy_paste_handler(copy, TeklaHandler.getInputData, error_text);
            copy_paste_handler(paste, TeklaHandler.getOutputData, error_text);
        }
        //private void btn_PasteSectionDetail_Click(object sender, EventArgs e)
        //{
        //    string error_text = "[ERROR] pasting";
        //    copy_paste_handler(paste, TeklaHandler.getOutputData, error_text);
        //}
        private void copy(Func<_ViewData> getter)
        {
            add_text("Copy... ");
            input = getter();
            add_text(input.countObjects());
            add_text("Done");
        }
        private void paste(Func<_ViewData> getter)
        {
        Begin:
            add_text("Paste... ");
            output = getter();
            TeklaHandler.copyView(input, output);
            add_text("Done");
            goto Begin;
        }
        private void copy_paste_handler(Action<Func<_ViewData>> function, Func<_ViewData> getter, string error)
        {
            toggle_all_controls(false);

            try
            {
                function(getter);
            }
            catch (OperationCanceledException)
            {
                add_text("[ERROR] Must select a point in a VIEW");
            }
            catch
            {
                add_text(error);
            }

            toggle_all_controls(true);
        }
        public void add_text(string message)
        {

        }
        public void replace_text(string message)
        {
            try
            {

            }
            catch
            {
                add_text(message);
            }
        }
        private void toggle_all_controls(bool status)
        {

        }

        private void btn_CreateSection_Click(object sender, EventArgs e)
        {
            //if (Control.ModifierKeys == Keys.Shift)
            //{
            //    MessageBox.Show("aaa");
            //    tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            //    tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            //}
            //try
            //{
            Function function = new Function();
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            List<tsd.Part> listPart = new List<tsd.Part>(); // danh sách part để cắt section
            List<tsd.SectionMarkBase> listSectionMarkBase = new List<tsd.SectionMarkBase>(); // danh sách section mark để đánh số
            foreach (tsd.DrawingObject item in drobjenum)
            {
                if (item is tsd.Part)
                {
                    tsd.Part part = item as tsd.Part;
                    listPart.Add(part);
                }
                else if (item is tsd.SectionMarkBase)
                {
                    tsd.SectionMarkBase sectionMarkBase = item as tsd.SectionMarkBase;
                    listSectionMarkBase.Add(sectionMarkBase);
                }
            }
            if (listPart.Count() != 0 && listSectionMarkBase.Count() == 0)
            #region Khi các đối tượng được chọn trên bản vẽ là tsd.Part thì sẽ tạo section
            {
                PartClassification partClassification = new PartClassification(listPart);
                List<tsm.Part> platesLeft = partClassification.PlatesLeft;
                List<tsm.Part> platesRight = partClassification.PlatesRight;
                List<tsm.Part> hotrollParts = partClassification.HotRollPart;
                tsd.View view = partClassification.View;
                double mainviewScale = view.Attributes.Scale;

                //Tạo ra danh sách part trái va phải
                List<tsm.Part> partLefts = new List<tsm.Part>();
                List<tsm.Part> partRights = new List<tsm.Part>();
                //Sap xep các cấu kiện theo phương X tăng dần
                partLefts = platesLeft.OrderBy(p => (new Part_Edge(view, p)).PointXmin0.X).ToList();
                partRights = platesRight.OrderBy(p => (new Part_Edge(view, p)).PointXmin0.X).ToList();

                List<tsm.Part> listPartLeftSec = new List<tsm.Part>(); // list part tạo section chỉ có part Left
                List<tsm.Part> listPartRightSec = new List<tsm.Part>();// list part tạo section chỉ có part Right
                List<tsm.Part> listPartBothSec = new List<tsm.Part>();// list part tạo section có cả part left, right
                List<GroupNSFS> listGroupNSFS = new List<GroupNSFS>();

                #region ĐOẠN NÀY CƯỜNG LÀM để lấy ra các nhóm part sẽ cắt section
                foreach (tsm.Part partLeft in partLefts)
                {
                    int count = 0;
                    foreach (tsm.Part partRight in partRights)
                    {
                        double KC = Math.Abs((new Part_Edge(view, partLeft).PointXmin0.X) - (new Part_Edge(view, partRight).PointXmin0.X));
                        if (KC <= 100)
                        {
                            listGroupNSFS.Add(new GroupNSFS(partLeft, partRight));
                            count++;
                        }
                    }
                    if (count == 0)
                        listPartLeftSec.Add(partLeft);
                }
                foreach (tsm.Part partRight in partRights)
                {
                    int count = 0;
                    foreach (tsm.Part partLeft in partLefts)
                    {
                        if (Math.Abs((new Part_Edge(view, partLeft).PointXmin0.X) - (new Part_Edge(view, partRight).PointXmin0.X)) <= 100)
                        {
                            listGroupNSFS.Add(new GroupNSFS(partLeft, partRight));
                            count++;
                        }
                    }
                    if (count == 0)
                        listPartRightSec.Add(partRight);
                }
                List<GroupNSFS> listGroupNSFSkhongtrung = new List<GroupNSFS>();
                foreach (GroupNSFS group in listGroupNSFS)
                {
                    List<GroupNSFS> exiting = listGroupNSFSkhongtrung.FindAll(p => p.PartMarkLeft == group.PartMarkLeft && p.PartMarkRight == group.PartMarkRight);
                    if (exiting.Count == 0)
                        listGroupNSFSkhongtrung.Add(group);
                }
                foreach (GroupNSFS group in listGroupNSFSkhongtrung)
                {
                    listPartBothSec.Add(group.PartRight);
                }
                List<tsm.Part> listPartLeftSecKhongTrung = new List<tsm.Part>(); // list part tạo section chỉ có part Left.
                List<tsm.Part> listPartRightSecKhongTrung = new List<tsm.Part>(); // list part tạo section chỉ có part Left.
                foreach (var item in listPartLeftSec)
                {
                    List<tsm.Part> existing = listPartLeftSecKhongTrung.FindAll(p => function.GetPartPos(p) == function.GetPartPos(item));
                    if (existing.Count == 0)
                    {
                        listPartLeftSecKhongTrung.Add(item);
                    }
                }
                foreach (var item in listPartRightSec)
                {
                    List<tsm.Part> existing = listPartRightSecKhongTrung.FindAll(p => function.GetPartPos(p) == function.GetPartPos(item));
                    if (existing.Count == 0)
                    {
                        listPartRightSecKhongTrung.Add(item);
                    }
                }

                listPartBothSec.AddRange(listPartLeftSecKhongTrung);
                listPartBothSec.AddRange(listPartRightSecKhongTrung);
                #endregion

                //while (drobjenum.MoveNext())
                //{
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                //Chuyen he toa do hien hanh ve he toa do cua View
                tsm.TransformationPlane viewPlane = new TransformationPlane(view.DisplayCoordinateSystem);
                wph.SetCurrentTransformationPlane(viewPlane);
                //model.CommitChanges();
                t3d.Point insertPoint = new t3d.Point();
                List<tsd.View> secViewList = new List<tsd.View>();
                //Duyệt danh sách các plate cần cắt section
                foreach (tsm.Part part in listPartBothSec) // list part tạo section có cả part left, right
                {
                    // MessageBox.Show(part.GetPartMark());
                    double thickness = function.GetWidth(part);
                    Part_Edge partEdge = new Part_Edge(view, part);
                    t3d.Point pointXminYmin = partEdge.PointXminYmin; // góc dưới bên trái
                    t3d.Point pointXminYmax = partEdge.PointXminYmax; // góc trên bên trái
                    t3d.Point pointXmaxYmin = partEdge.PointXmaxYmin; // góc dưới bên phải
                    t3d.Point pointXmaxYmax = partEdge.PointXmaxYmax; // góc trên bên phải
                    t3d.Point maxX = partEdge.PointXmax0;
                    t3d.Point minX = partEdge.PointXmin0;
                    t3d.Point maxY = partEdge.PointYmax0;
                    t3d.Point minY = partEdge.PointYmin0;

                    tsd.View secView = null;
                    tsd.SectionMark secMark = null;
                    tsd.SectionMarkBase.SectionMarkAttributes sectionMarkAttributes = new SectionMarkBase.SectionMarkAttributes();
                    sectionMarkAttributes.LoadAttributes(cb_SecMarkAttribute.Text);
                    tsd.View.ViewAttributes viewAttributes = new tsd.View.ViewAttributes();

                    viewAttributes.LoadAttributes(cb_SecAttribute.Text);
                    if (chbox_SameScale.Checked)
                        viewAttributes.Scale = mainviewScale;
                    //Tấm đứng
                    if (Math.Abs(maxX.X - minX.X - thickness) < 3)
                    {
                        //function.InsertSymbol(view, minX);
                        t3d.Point secSp = new t3d.Point();
                        t3d.Point secEp = new t3d.Point();
                        if (imcb_SectionSide.SelectedIndex == 0 || imcb_SectionSide.SelectedIndex == -1) //left top
                        {
                            secSp = pointXminYmin;
                            secEp = pointXminYmax;
                        }
                        else if (imcb_SectionSide.SelectedIndex == 1) //right bot
                        {
                            secSp = pointXmaxYmin;
                            secEp = pointXmaxYmax;
                        }

                        //function.InsertSymbol(view, secSp);
                        //function.InsertSymbol(view, secEp);
                        t3d.Vector vx = new Vector(secSp - secEp);
                        vx.Normalize(50); //Tính từ điểm secSp lên 50
                        secSp.Translate(vx.X, vx.Y, vx.Z);
                        vx = new Vector(secEp - secSp);
                        vx.Normalize(50);//Tính từ điểm secEp xuống 50
                        secEp.Translate(vx.X, vx.Y, vx.Z);
                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            //cắt từ điểm thấp đến điểm cao: Cắt từ Phải nhìn sang Trái
                            tsd.View.CreateSectionView(view, secSp, secEp, insertPoint, thickness + 50, 40, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                        else
                        {
                            //cắt từ điểm cao đến điểm thấp: Cắt từ Trái nhìn sang Phải
                            tsd.View.CreateSectionView(view, secEp, secSp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                    }
                    //Cheo phai
                    else if (Math.Abs(t3d.Distance.PointToPoint(minX, minY) - thickness) < 3 && Math.Abs(t3d.Distance.PointToPoint(maxX, maxY) - thickness) < 3)
                    {
                        t3d.Point secSp = new t3d.Point();
                        t3d.Point secEp = new t3d.Point();
                        if (imcb_SectionSide.SelectedIndex == 0 || imcb_SectionSide.SelectedIndex == -1) //left top
                        {
                            secSp = new t3d.Point(minX.X, minX.Y);
                            secEp = new t3d.Point(maxY.X, maxY.Y);
                        }
                        else if (imcb_SectionSide.SelectedIndex == 1) //right bot
                        {
                            secSp = new t3d.Point(minY.X, minY.Y);
                            secEp = new t3d.Point(maxX.X, maxX.Y);
                        }
                        t3d.Vector vx = new Vector(secSp - secEp);
                        vx.Normalize(50);
                        secSp.Translate(vx.X, vx.Y, vx.Z);
                        vx = new Vector(secEp - secSp);
                        vx.Normalize(50);
                        secEp.Translate(vx.X, vx.Y, vx.Z);

                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            //cắt từ điểm thấp đến điểm cao: Cắt từ Phải nhìn sang Trái
                            tsd.View.CreateSectionView(view, secSp, secEp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                        else
                        {
                            //cắt từ điểm cao đến điểm thấp: Cắt từ Trái nhìn sang Phải
                            tsd.View.CreateSectionView(view, secEp, secSp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                    }
                    //Cheo trái
                    else if (Math.Abs(t3d.Distance.PointToPoint(maxX, minY) - thickness) < 3 && Math.Abs(t3d.Distance.PointToPoint(maxY, minX) - thickness) < 3)
                    {
                        t3d.Point secSp = new t3d.Point();
                        t3d.Point secEp = new t3d.Point();
                        if (imcb_SectionSide.SelectedIndex == 0 || imcb_SectionSide.SelectedIndex == -1) //left top
                        {
                            secSp = new t3d.Point(minX.X, minX.Y);
                            secEp = new t3d.Point(minY.X, minY.Y);
                        }
                        else if (imcb_SectionSide.SelectedIndex == 1) //right bot
                        {
                            secSp = new t3d.Point(maxY.X, maxY.Y);
                            secEp = new t3d.Point(maxX.X, maxX.Y);
                        }
                        t3d.Vector vx = new Vector(secSp - secEp);
                        vx.Normalize(50);
                        secSp.Translate(vx.X, vx.Y, vx.Z);
                        vx = new Vector(secEp - secSp);
                        vx.Normalize(50);
                        secEp.Translate(vx.X, vx.Y, vx.Z);
                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            //cắt từ điểm thấp đến điểm cao: Cắt từ Phải nhìn sang Trái
                            tsd.View.CreateSectionView(view, secSp, secEp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                        else
                        {
                            //cắt từ điểm cao đến điểm thấp: Cắt từ Trái nhìn sang Phải
                            tsd.View.CreateSectionView(view, secEp, secSp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                    }
                    //Tấm ngang
                    else if (Math.Abs(maxY.Y - minY.Y - thickness) < 3)
                    {
                        t3d.Point secSp = new t3d.Point();
                        t3d.Point secEp = new t3d.Point();
                        if (imcb_SectionSide.SelectedIndex == 0 || imcb_SectionSide.SelectedIndex == -1) //left top
                        {
                            secSp = pointXminYmax;
                            secEp = pointXmaxYmax;
                        }
                        else if (imcb_SectionSide.SelectedIndex == 1) //right bot
                        {
                            secSp = pointXminYmin;
                            secEp = pointXmaxYmin;
                        }
                        t3d.Vector vx = new Vector(secSp - secEp);
                        vx.Normalize(50);
                        secSp.Translate(vx.X, vx.Y, vx.Z);
                        vx = new Vector(secEp - secSp);
                        vx.Normalize(50);
                        secEp.Translate(vx.X, vx.Y, vx.Z);
                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            //cắt từ điểm Trái qua: Cắt từ Duói nhìn lên
                            tsd.View.CreateSectionView(view, secSp, secEp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                        else
                        {
                            //cắt từ điểm Phải qua: Cắt từ trên nhìn xuống
                            tsd.View.CreateSectionView(view, secEp, secSp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                    }
                    //Trường hợp còn lại sẽ cắt từ trên nhìn xuống
                    else
                    {
                        t3d.Point secSp = pointXminYmax;
                        t3d.Point secEp = pointXmaxYmax;
                        t3d.Vector vx = new Vector(secSp - secEp);
                        vx.Normalize(50);
                        secSp.Translate(vx.X, vx.Y, vx.Z);
                        vx = new Vector(secEp - secSp);
                        vx.Normalize(50);
                        secEp.Translate(vx.X, vx.Y, vx.Z);
                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            //cắt từ điểm Trái qua: Cắt từ Duói nhìn lên
                            tsd.View.CreateSectionView(view, secSp, secEp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                        else
                        {
                            //cắt từ điểm Phải qua: Cắt từ trên nhìn xuống
                            tsd.View.CreateSectionView(view, secEp, secSp, insertPoint, thickness + 50, 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                        }
                    }
                    secViewList.Add(secView);
                }
                //Duyệt danh sách các hotrollpart cần cắt section
                foreach (tsm.Part part in hotrollParts)
                {
                    // MessageBox.Show(part.GetPartMark());
                    ArrayList listCenterPoint = part.GetCenterLine(true);
                    double partHeight = function.GetHeight(part);
                    t3d.Point secSp = listCenterPoint[0] as t3d.Point;
                    t3d.Point secEp = listCenterPoint[1] as t3d.Point;
                    tsd.View secView = null;
                    tsd.SectionMark secMark = null;
                    tsd.SectionMarkBase.SectionMarkAttributes sectionMarkAttributes = new SectionMarkBase.SectionMarkAttributes();
                    sectionMarkAttributes.LoadAttributes(cb_SecMarkAttribute.Text);
                    tsd.View.ViewAttributes viewAttributes = new tsd.View.ViewAttributes();
                    viewAttributes.LoadAttributes(cb_SecAttribute.Text);
                    t3d.Vector vx = new Vector(secSp - secEp);
                    vx.Normalize(50); //Tính từ điểm secSp lên 50
                    secSp.Translate(vx.X, vx.Y, vx.Z);
                    vx = new Vector(secEp - secSp);
                    vx.Normalize(50);//Tính từ điểm secEp xuống 50
                    secEp.Translate(vx.X, vx.Y, vx.Z);
                    if (Control.ModifierKeys == Keys.Shift)
                    {
                        //cắt từ điểm thấp đến điểm cao: Cắt từ Phải nhìn sang Trái
                        tsd.View.CreateSectionView(view, secSp, secEp, insertPoint, partHeight / 2 + 50, partHeight + 40, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                    }
                    else
                    {
                        //cắt từ điểm cao đến điểm thấp: Cắt từ Trái nhìn sang Phải
                        tsd.View.CreateSectionView(view, secEp, secSp, insertPoint, partHeight / 2 + 50, partHeight + 50, viewAttributes, sectionMarkAttributes, out secView, out secMark);
                    }
                }
                for (int i = 0; i < secViewList.Count; i++)
                {
                    if (i == 0)
                    {
                        insertPoint = new t3d.Point(secViewList[i].Width, 0, 0);
                        continue;
                    }
                    insertPoint = new t3d.Point(insertPoint.X + secViewList[i].Width, 0, 0);
                    secViewList[i].Origin = insertPoint;
                    if (secViewList[i].Attributes.Scale < view.Attributes.Scale)
                        secViewList[i].Attributes.Scale = view.Attributes.Scale;
                    secViewList[i].Modify();
                }
            }
            #endregion

            //Nếu section mark được chọn và không có part nào được chọn sẽ kích hoạt chức năng gióng thẳng hàng secion mark
            else if (listSectionMarkBase.Count() != 0 && listPart.Count() == 0)
            {
                if (Control.ModifierKeys != Keys.Shift)
                {
                Begin:
                    tsdui.Picker c_picker = drawingHandler.GetPicker();
                    var pick = c_picker.PickPoint("Pick first point");
                    t3d.Point point1 = pick.Item1;
                    tsd.View view = pick.Item2 as tsd.View;
                    double scale = view.Attributes.Scale;
                    t3d.Point point2 = c_picker.PickPoint("Pick second point").Item1;
                    //Lấy view hiện hành thông qua điểm vừa pick
                    //tsm.TransformationPlane viewplane = new TransformationPlane(view.DisplayCoordinateSystem);
                    //wph.SetCurrentTransformationPlane(viewplane);
                    foreach (SectionMarkBase sectionMarkBase in listSectionMarkBase)
                    {
                        if (Math.Abs(sectionMarkBase.LeftPoint.X - sectionMarkBase.RightPoint.X) < 5 && Math.Abs(point1.X - point2.X) < 5)
                        {
                            sectionMarkBase.Attributes.LineWidthOffsetLeft = Math.Abs(point1.Y - sectionMarkBase.LeftPoint.Y) * 1 / scale;
                            sectionMarkBase.Attributes.LineWidthOffsetRight = Math.Abs(point2.Y - sectionMarkBase.RightPoint.Y) * 1 / scale;
                            sectionMarkBase.Modify();
                        }
                        else if (Math.Abs(sectionMarkBase.LeftPoint.Y - sectionMarkBase.RightPoint.Y) < 5 && Math.Abs(point1.Y - point2.Y) < 5)

                        {
                            sectionMarkBase.Attributes.LineWidthOffsetLeft = Math.Abs(point2.X - sectionMarkBase.LeftPoint.X) * 1 / scale;
                            sectionMarkBase.Attributes.LineWidthOffsetRight = Math.Abs(point1.X - sectionMarkBase.RightPoint.X) * 1 / scale;
                            sectionMarkBase.Modify();
                        }
                    }
                    goto Begin;
                }

                else if (Control.ModifierKeys == Keys.Shift)
                {
                    //string sectionMarkPrefix = tb_SectionMarkPrefix.Text.Trim(); //bỏ dấu trắng nếu có
                    //string sectionMarkNo = tb_SectionMarkNo.Text.Trim(); //bỏ dấu trắng nếu có
                    //char[] alphabet = Enumerable.Range('A', 26).Select(x => (char)x).ToArray();
                    //listSectionMarkBase = listSectionMarkBase.OrderBy(p => (p.LeftPoint.X)).ToList();

                    //if (int.TryParse(sectionMarkNo, out int sectionMarkNoInt)) //Nếu số tự tự là 1 số
                    //{
                    //    int.TryParse(sectionMarkNo, out sectionMarkNoInt);
                    //    foreach (SectionMarkBase sectionMarkBase in listSectionMarkBase)
                    //    {
                    //        sectionMarkBase.Attributes.MarkName = sectionMarkPrefix + sectionMarkNoInt.ToString();
                    //        sectionMarkBase.Modify();
                    //        sectionMarkNoInt += 1;
                    //    }
                    //}
                    //else if (char.TryParse(sectionMarkNo, out char sectionMarkNoChar))// Nếu số thứ tự là A B C D...
                    //{
                    //    sectionMarkNoChar = char.Parse(sectionMarkNo);
                    //    int charIndex = Array.IndexOf(alphabet, sectionMarkNoChar); // lấy thứ tự của kí tự trong ô number ( B hoặc C D ...)
                    //    foreach (SectionMarkBase sectionMarkBase in listSectionMarkBase)
                    //    {
                    //        sectionMarkBase.Attributes.MarkName = sectionMarkPrefix + alphabet[charIndex].ToString();
                    //        sectionMarkBase.Modify();
                    //        charIndex += 1;
                    //    }
                    //}
                }
            }
            else if (listSectionMarkBase.Count() != 0 && listPart.Count() != 0)
                MessageBox.Show("Please select Part or SectionMark only!");
            drawingHandler.GetActiveDrawing().CommitChanges();
            //}
            //catch
            //{ }

            #region OLD CODE
            //if (drobjenum.Current is tsd.Part)
            //{
            //    tsd.Part dr_part = drobjenum.Current as tsd.Part;
            //    tsm.Part m_part = model.SelectModelObject(dr_part.ModelIdentifier) as tsm.Part;
            //    tsd.View viewcur = dr_part.GetView() as tsd.View;
            //    tsd.ViewBase viewbase = dr_part.GetView();
            //    tsm.TransformationPlane viewplane = new TransformationPlane(viewcur.DisplayCoordinateSystem);
            //    wph.SetCurrentTransformationPlane(viewplane);
            //    Part_Edge part_edge = new Part_Edge(viewcur, m_part);
            //    List<t3d.Point> Part_PL = part_edge.List_Edge;
            //    t3d.Point minX = Part_PL.OrderBy(p => p.X).ToList()[0];
            //    t3d.Point maxX = Part_PL.OrderByDescending(p => p.X).ToList()[0];
            //    t3d.Point minY = Part_PL.OrderBy(p => p.Y).ToList()[0];
            //    t3d.Point maxY = Part_PL.OrderByDescending(p => p.Y).ToList()[0];
            //    double cday = 0;
            //    m_part.GetReportProperty("PROFILE.WIDTH", ref cday);

            //    if (cbx_Reverse.Checked == false)
            //    {
            //        if (Math.Abs(maxX.X - minX.X - cday) < 3) //Tấm đứng
            //        {
            //            t3d.Point p_T = new t3d.Point(maxX.X, maxY.Y + 20); //Tạo điểm phía trên
            //            t3d.Point p_D = new t3d.Point(maxX.X, minY.Y - 20);//Tạo điểm phía dưới
            //            tsd.View.CreateSectionView(viewcur, p_D, p_T, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //        else if (Math.Abs(t3d.Distance.PointToPoint(minX, minY) - cday) < 3)// Tấm nghiêng phải
            //        {
            //            tsd.View.CreateSectionView(viewcur, minY, maxX, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //        else if (Math.Abs(t3d.Distance.PointToPoint(minY, maxX) - cday) < 3)// Tấm nghiêng trái
            //        {
            //            tsd.View.CreateSectionView(viewcur, minY, minX, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //        if (Math.Abs(maxY.Y - minY.Y - cday) < 3) //Tấm ngang
            //        {
            //            t3d.Point p_PH = new t3d.Point(maxX.X + 20, maxY.Y); //Tạo điểm phía trên bên phải
            //            t3d.Point p_TR = new t3d.Point(minX.X - 20, maxY.Y);//Tạo điểm phía dưới
            //            tsd.View.CreateSectionView(viewcur, p_PH, p_TR, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //    }
            //    else
            //    {
            //        if (Math.Abs(maxX.X - minX.X - cday) < 3) //Tấm đứng
            //        {
            //            t3d.Point p_T = new t3d.Point(maxX.X, maxY.Y + 20); //Tạo điểm phía trên
            //            t3d.Point p_D = new t3d.Point(maxX.X, minY.Y - 20);//Tạo điểm phía dưới
            //            tsd.View.CreateSectionView(viewcur, p_T, p_D, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //        else if (Math.Abs(t3d.Distance.PointToPoint(minX, minY) - cday) < 3)// Tấm nghiêng phải
            //        {
            //            tsd.View.CreateSectionView(viewcur, maxX, minY, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //        else if (Math.Abs(t3d.Distance.PointToPoint(minY, maxX) - cday) < 3)// Tấm nghiêng trái
            //        {
            //            tsd.View.CreateSectionView(viewcur, minX, minY, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //        if (Math.Abs(maxY.Y - minY.Y - cday) < 3) //Tấm ngang
            //        {
            //            t3d.Point p_PH = new t3d.Point(maxX.X + 20, maxY.Y); //Tạo điểm phía trên bên phải
            //            t3d.Point p_TR = new t3d.Point(minX.X - 20, maxY.Y);//Tạo điểm phía dưới
            //            tsd.View.CreateSectionView(viewcur, p_TR, p_PH, new t3d.Point(100 + diemdat, 100), 50, 50, viewattr, sec_label, out sectionview, out sec_mark);
            //        }
            //    }
            //    diemdat = diemdat + 50;
            //}
            //}
            #endregion
            wph.SetCurrentTransformationPlane(new TransformationPlane());
            model.CommitChanges();
        }

        public void Save_View_Attribure_Temporary()
        {
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                            "{" +
                            " public class Script" +
                                "{" +
                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                    " {" +
            @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
            @"akit.ValueChange(""view_dial"", ""gr_view_saveas_file"", ""temporary_view_save"";" +
            @"akit.PushButton(""gr_view_saveas"", ""view_dial"");" +
            @" akit.PushButton(""view_ok"", ""view_dial"");" +
            "}}}";
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }

        private void btn_NumberingSectionMark_Click(object sender, EventArgs e)
        {
            Function function = new Function();
            tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
            DrawingObjectEnumerator.AutoFetch = true;
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            List<tsd.SectionMarkBase> listSectionMarkBase = new List<tsd.SectionMarkBase>(); // danh sách section mark để đánh số
            ArrayList selectedMarks = new ArrayList();
            ArrayList finalSelectedMarks = new ArrayList();
            foreach (tsd.DrawingObject item in drobjenum)
            {
                tsd.SectionMarkBase sectionMarkBase = item as tsd.SectionMarkBase;
                listSectionMarkBase.Add(sectionMarkBase);
                finalSelectedMarks.Add(sectionMarkBase);
            }
            drawingHandler.GetDrawingObjectSelector().UnselectAllObjects();
            string sectionMarkPrefix = tb_SectionMarkPrefix.Text.Trim(); //bỏ dấu trắng nếu có
            string sectionMarkNo = tb_SectionMarkNo.Text.Trim(); //bỏ dấu trắng nếu có
            char[] alphabet = Enumerable.Range('A', 26).Select(x => (char)x).ToArray();
            listSectionMarkBase = listSectionMarkBase.OrderBy(p => (p.LeftPoint.X)).ToList();
            if (Control.ModifierKeys == Keys.Shift)
                listSectionMarkBase = listSectionMarkBase.OrderByDescending(p => (p.LeftPoint.X)).ToList();
            if (int.TryParse(sectionMarkNo, out int sectionMarkNoInt)) //Nếu số tự tự là 1 số
            {
                int.TryParse(sectionMarkNo, out sectionMarkNoInt);
                foreach (SectionMarkBase sectionMarkBase in listSectionMarkBase)
                {
                    string sectionMarkName = sectionMarkPrefix + sectionMarkNoInt.ToString();
                    drawingHandler.GetDrawingObjectSelector().UnselectAllObjects();
                    selectedMarks.Add(sectionMarkBase);
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(selectedMarks, false);
                    RunMacroChangeSectionName(sectionMarkName);
                    selectedMarks.Clear();
                    sectionMarkNoInt += 1;
                }
            }
            else if (char.TryParse(sectionMarkNo, out char sectionMarkNoChar))// Nếu số thứ tự là A B C D...
            {
                sectionMarkNoChar = char.Parse(sectionMarkNo);
                int charIndex = Array.IndexOf(alphabet, sectionMarkNoChar); // lấy thứ tự của kí tự trong ô number ( B hoặc C D ...)
                foreach (SectionMarkBase sectionMarkBase in listSectionMarkBase)
                {
                    string sectionMarkName = sectionMarkPrefix + alphabet[charIndex].ToString();
                    drawingHandler.GetDrawingObjectSelector().UnselectAllObjects();
                    selectedMarks.Add(sectionMarkBase);
                    drawingHandler.GetDrawingObjectSelector().SelectObjects(selectedMarks, false);
                    RunMacroChangeSectionName(sectionMarkName);
                    selectedMarks.Clear();
                    charIndex += 1;
                }
            }
            drawingHandler.GetDrawingObjectSelector().SelectObjects(finalSelectedMarks, false);
        }

        public void RunMacroChangeSectionName(string markText)
        {
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }

            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                            "{" +
                            " public class Script" +
                                "{" +
                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                    " {" +
            @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
            @"akit.PushButton(""csym_on_off"", ""csym_dial"");" +
            @"akit.ValueChange(""csym_dial"", ""csym_label"", """ + markText + @""");" +
            @"akit.PushButton(""csym_modify"", ""csym_dial"");" +
            @" akit.PushButton(""csym_cancel"", ""csym_dial"");" +
            "}}}";
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);

        }

        #region Code của Drawing Tool 2
        private void tb_DrawingName_Enter(object sender, EventArgs e)
        {
            timer1.Enabled = false;
        }// dùng chung cho tb_Title1,2,3_Enter luôn

        private void tb_DrawingName_Leave(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (tb_DrawingName.Modified)
                {
                    cur_dr.Name = tb_DrawingName.Text;
                    cur_dr.Modify();
                    cur_dr.CommitChanges();
                }
                timer1.Enabled = true;
            }
            catch
            { }
        }

        private void tb_DrawingTitle1_Leave(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (tb_DrawingTitle1.Modified)
                {
                    cur_dr.Title1 = tb_DrawingTitle1.Text;
                    cur_dr.Modify();
                    cur_dr.CommitChanges();
                }
                timer1.Enabled = true;
            }
            catch (Exception)
            { }
        }

        private void tb_DrawingTitle2_Leave(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (tb_DrawingTitle2.Modified)
                {
                    cur_dr.Title2 = tb_DrawingTitle2.Text;
                    cur_dr.Modify();
                    cur_dr.CommitChanges();
                }
                timer1.Enabled = true;
            }
            catch (Exception)
            { }
        }

        private void tb_DrawingTitle3_Leave(object sender, EventArgs e)
        {
            try
            {
                timer1.Enabled = false;
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (tb_DrawingTitle3.Modified)
                {
                    cur_dr.Title3 = tb_DrawingTitle3.Text;
                    cur_dr.Modify();
                }
                timer1.Enabled = true;
            }
            catch
            { }
        }

        private void btn_Frozen_Click(object sender, EventArgs e)
        {
            try
            {
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                if (cur_dr.IsFrozen == false)
                {
                    cur_dr.IsFrozen = true;
                    btn_Frozen.BackColor = System.Drawing.Color.DeepSkyBlue;
                }
                else
                {
                    cur_dr.IsFrozen = false;
                    btn_Frozen.BackColor = System.Drawing.Color.Silver;
                }
                cur_dr.Modify();
                if (chb_SaveAndNextDrawingFbutton.Checked)
                {
                    drawingHandler.SaveActiveDrawing();
                    OpenNextDrawing();
                }
            }
            catch
            { }
        }

        public void OpenNextDrawing()
        {
            //cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }

            //if (cur_dr is tsd.Drawing)
            //{
            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                            "{" +
                            " public class Script" +
                                "{" +
                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                    " {" +
            @"akit.Callback(""grOpenNextDrawingCB"", """", ""main_frame"");" +
            "}}}";
            //}

            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }

        public void OpenPreviousDrawing()
        {
            //cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }

            //if (cur_dr is tsd.Drawing)
            //{
            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                            "{" +
                            " public class Script" +
                                "{" +
                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                    " {" +
            @"akit.Callback(""grOpenPreviousDrawingCB"", """", ""main_frame"");" +
            "}}}";
            //}

            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }
        private void btn_Issue_Click(object sender, EventArgs e)
        {
            try
            {
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                //Lấy bản vẽ được select trong drawing list
                DrawingEnumerator selectedDrawings = drawingHandler.GetDrawingSelector().GetSelected();
                while (selectedDrawings.MoveNext())
                {
                    tsd.Drawing selected_dr = selectedDrawings.Current;
                    if (selected_dr.Mark == cur_dr.Mark)
                    {
                        string Name = "TemporaryMacro.cs";//Temporary file name.
                        string MacrosPath = string.Empty;//would store Path to script
                        string script2 = "";
                        Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
                        if (cur_dr.IsIssued == false)
                        {
                            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
                            //Create a script as string
                            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                            "{" +
                                            " public class Script" +
                                                "{" +
                                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                                    " {" +
                            @"akit.PushButton(""dia_draw_issue_on"", ""Drawing_selection"");" +
                            "}}}";

                            //Save to file
                            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
                            //Run it at Tekla Structures application
                            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
                            btn_Issue.BackColor = System.Drawing.Color.DarkOrange;
                        }
                        else
                        {
                            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
                            //Create a script as string
                            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                            "{" +
                                            " public class Script" +
                                                "{" +
                                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                                    " {" +
                            @"akit.PushButton(""dia_draw_issue_off"", ""Drawing_selection"");" +
                            "}}}";

                            //Save to file
                            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
                            //Run it at Tekla Structures application
                            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
                            btn_Issue.BackColor = System.Drawing.Color.Silver;
                        }
                        cur_dr.Modify();
                    }
                    else MessageBox.Show("Active drawing has not been selected in the drawing list.");
                }
                if (chb_SaveAndNextDrawingFbutton.Checked)
                {
                    drawingHandler.SaveActiveDrawing();
                    OpenNextDrawing();
                }
            }
            catch
            { }
        }

        private void btn_Ready_Click(object sender, EventArgs e)
        {
            try
            {
                tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
                //Lấy bản vẽ được select trong drawing list
                DrawingEnumerator selectedDrawings = drawingHandler.GetDrawingSelector().GetSelected();
                while (selectedDrawings.MoveNext())
                {
                    tsd.Drawing selected_dr = selectedDrawings.Current;
                    if (selected_dr.Mark == cur_dr.Mark)
                    {
                        string Name = "TemporaryMacro.cs";//Temporary file name.
                        string MacrosPath = string.Empty;//would store Path to script
                        string script2 = "";
                        Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
                        if (cur_dr.IsReadyForIssue == false)
                        {
                            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
                            //Create a script as string
                            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                            "{" +
                                            " public class Script" +
                                                "{" +
                                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                                    " {" +
                            @"akit.PushButton(""dia_draw_ready_for_issue_on"", ""Drawing_selection"");" +
                            "}}}";

                            //Save to file
                            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
                            //Run it at Tekla Structures application
                            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);

                            btn_Ready.BackColor = System.Drawing.Color.Green;

                        }
                        else
                        {
                            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
                            //Create a script as string
                            script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                            "{" +
                                            " public class Script" +
                                                "{" +
                                                    " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                                    " {" +
                            @"akit.PushButton(""dia_draw_ready_for_issue_off"", ""Drawing_selection"");" +
                            "}}}";

                            //Save to file
                            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
                            //Run it at Tekla Structures application
                            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);

                            btn_Ready.BackColor = System.Drawing.Color.Silver;
                        }
                        // cur_dr.Modify();  // vì bản vẽ sẽ mất dấu tích xanh ready for issue khi bản vẽ bị modify, nên phải bỏ dòng này đi.
                    }
                    else MessageBox.Show("Active drawing has not been selected in the drawing list.");
                }
                if (chb_SaveAndNextDrawingFbutton.Checked)
                {
                    drawingHandler.SaveActiveDrawing();
                    OpenNextDrawing();
                }
            }
            catch
            { }
        }

        private void btn_ScaleChange1_Click(object sender, EventArgs e)
        {
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drobjenum.GetSize() == 0)
                ChangeDrawingScale(sender);
            else
            {
                foreach (tsd.View view in drobjenum)
                {
                    view.Attributes.Scale = Convert.ToDouble((sender as Button).Text);
                    view.Modify();
                }
            }
        }
        private void btn_ChangeScale_Click(object sender, EventArgs e)
        {
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            string scale = tb_CurrentScale.Text;
            if (drobjenum.GetSize() == 0)
                ChangeDrawingScaleOnly(scale);
            else
            {
                foreach (tsd.View view in drobjenum)
                {
                    view.Attributes.Scale = Convert.ToDouble(scale);
                    view.Modify();
                }
            }
        }
        public void ChangeDrawingScale(object sender)
        {
            tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
            double minCutLength = new double();
            try
            {
                minCutLength = Convert.ToDouble(tb_MinimumCutPartLength.Text);
            }
            catch
            { }
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
            if (cur_dr is tsd.AssemblyDrawing)
            {
                if (tb_MinimumCutPartLength.Text == "")
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""gr_adraw_on_off"", ""adraw_dial"");" +
                    @"akit.PushButton(""gr_adraw_view"", ""adraw_dial"");" +
                    @"akit.PushButton(""dv_on_off"", ""adv_dial"");" +
                    @"akit.TabChange(""adv_dial"", ""contMain"", ""tabAttributes"");" +
                    @"akit.ValueChange(""adv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                    @"akit.PushButton(""dv_modify"", ""adv_dial"");" +
                    @" akit.PushButton(""gr_adraw_ok"", ""adraw_dial"");" +
                    "}}}";
                }
                else
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""gr_adraw_on_off"", ""adraw_dial"");" +
                    @"akit.PushButton(""gr_adraw_view"", ""adraw_dial"");" +
                    @"akit.PushButton(""dv_on_off"", ""adv_dial"");" +
                    @"akit.TabChange(""adv_dial"", ""contMain"", ""tabAttributes"");" +
                    @"akit.ValueChange(""adv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                    @"akit.TabChange(""adv_dial"", ""contMain"", ""tabShortening"");" +
                    @"akit.ValueChange(""adv_dial"",""gr_view_cut_min_dist"",""" + minCutLength + @""");" +
                    @"akit.PushButton(""dv_modify"", ""adv_dial"");" +
                    @" akit.PushButton(""gr_adraw_ok"", ""adraw_dial"");" +
                    "}}}";
                }
            }
            else if (cur_dr is tsd.SinglePartDrawing)
            {
                if (tb_MinimumCutPartLength.Text == "")
                {  //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                "{" +
                                " public class Script" +
                                    "{" +
                                        " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                        " {" +
                @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                @"akit.PushButton(""gr_wdraw_on_off"", ""wdraw_dial"");" +
                @"akit.PushButton(""gr_wdraw_view"", ""wdraw_dial"");" +
                @"akit.PushButton(""dv_on_off"", ""wdv_dial"");" +
                @"akit.TabChange(""wdv_dial"", ""contMain"", ""tabAttributes"");" +
                @"akit.ValueChange(""wdv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                 @"akit.PushButton(""dv_modify"", ""wdv_dial"");" +
                @" akit.PushButton(""gr_wdraw_ok"", ""wdraw_dial"");" +
                "}}}";
                }
                else
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                "{" +
                                " public class Script" +
                                    "{" +
                                        " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                        " {" +
                @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                @"akit.PushButton(""gr_wdraw_on_off"", ""wdraw_dial"");" +
                @"akit.PushButton(""gr_wdraw_view"", ""wdraw_dial"");" +
                @"akit.PushButton(""dv_on_off"", ""wdv_dial"");" +
                @"akit.TabChange(""wdv_dial"", ""contMain"", ""tabAttributes"");" +
                @"akit.ValueChange(""wdv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                @"akit.TabChange(""wdv_dial"", ""contMain"", ""tabShortening"");" +
                @"akit.ValueChange(""wdv_dial"",""gr_view_cut_min_dist"",""" + minCutLength + @""");" +
                @"akit.PushButton(""dv_modify"", ""wdv_dial"");" +
                @" akit.PushButton(""gr_wdraw_ok"", ""wdraw_dial"");" +
                "}}}";
                }

            }
            else if (cur_dr is tsd.GADrawing)
            {
                if (tb_MinimumCutPartLength.Text == "")
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""gr_gdraw_on_off"", ""gdraw_dial"");" +
                    @"akit.PushButton(""gr_gdraw_view"", ""gdraw_dial"");" +
                    @"akit.PushButton(""dv_on_off"", ""gdv_dial"");" +
                    @"akit.TabChange(""gdv_dial"", ""contMain"", ""tabAttributes"");" +
                    @"akit.ValueChange(""gdv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                    @"akit.PushButton(""dv_modify"", ""gdv_dial"");" +
                    @" akit.PushButton(""gr_gdraw_ok"", ""gdraw_dial"");" +
                    "}}}";
                }
                else
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""gr_gdraw_on_off"", ""gdraw_dial"");" +
                    @"akit.PushButton(""gr_gdraw_view"", ""gdraw_dial"");" +
                    @"akit.PushButton(""dv_on_off"", ""gdv_dial"");" +
                    @"akit.TabChange(""gdv_dial"", ""contMain"", ""tabAttributes"");" +
                    @"akit.ValueChange(""gdv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                    @"akit.TabChange(""gdv_dial"", ""contMain"", ""tabShortening"");" +
                    @"akit.ValueChange(""gdv_dial"",""gr_view_cut_min_dist"",""" + minCutLength + @""");" +
                    @"akit.PushButton(""dv_modify"", ""gdv_dial"");" +
                    @" akit.PushButton(""gr_gdraw_ok"", ""gdraw_dial"");" +
                    "}}}";
                }
            }
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }
        public void ChangeDrawingScaleOnly(string scale)
        {
            tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
            double minCutLength = new double();
            try
            {
                minCutLength = Convert.ToDouble(tb_MinimumCutPartLength.Text);
            }
            catch
            { }
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";

            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
            if (cur_dr is tsd.AssemblyDrawing)
            {
                //Create a script as string
                script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                "{" +
                                " public class Script" +
                                    "{" +
                                        " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                        " {" +
                @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                @"akit.PushButton(""gr_adraw_on_off"", ""adraw_dial"");" +
                @"akit.PushButton(""gr_adraw_view"", ""adraw_dial"");" +
                @"akit.PushButton(""dv_on_off"", ""adv_dial"");" +
                @"akit.TabChange(""adv_dial"", ""contMain"", ""tabAttributes"");" +
                @"akit.ValueChange(""adv_dial"", ""gr_dv_main_scale"", """ + scale + @""");" +
                @"akit.PushButton(""dv_modify"", ""adv_dial"");" +
                @" akit.PushButton(""gr_adraw_ok"", ""adraw_dial"");" +
                "}}}";
            }
            else if (cur_dr is tsd.SinglePartDrawing)
            {
                script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                "{" +
                                " public class Script" +
                                    "{" +
                                        " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                        " {" +
                @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                @"akit.PushButton(""gr_wdraw_on_off"", ""wdraw_dial"");" +
                @"akit.PushButton(""gr_wdraw_view"", ""wdraw_dial"");" +
                @"akit.PushButton(""dv_on_off"", ""wdv_dial"");" +
                @"akit.TabChange(""wdv_dial"", ""contMain"", ""tabAttributes"");" +
                @"akit.ValueChange(""wdv_dial"", ""gr_dv_main_scale"", """ + scale + @""");" +
                 @"akit.PushButton(""dv_modify"", ""wdv_dial"");" +
                @" akit.PushButton(""gr_wdraw_ok"", ""wdraw_dial"");" +
                "}}}";
            }
            else if (cur_dr is tsd.GADrawing)
            {
                //Create a script as string
                script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                "{" +
                                " public class Script" +
                                    "{" +
                                        " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                        " {" +
                @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                @"akit.PushButton(""gr_gdraw_on_off"", ""gdraw_dial"");" +
                @"akit.PushButton(""gr_gdraw_view"", ""gdraw_dial"");" +
                @"akit.PushButton(""dv_on_off"", ""gdv_dial"");" +
                @"akit.TabChange(""gdv_dial"", ""contMain"", ""tabAttributes"");" +
                @"akit.ValueChange(""gdv_dial"", ""gr_dv_main_scale"", """ + scale + @""");" +
                @"akit.PushButton(""dv_modify"", ""gdv_dial"");" +
                @" akit.PushButton(""gr_gdraw_ok"", ""gdraw_dial"");" +
                "}}}";
            }
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }

        private void tb_CurrentScale_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btn_ChangeScale_Click(this, new EventArgs());
        }
        private void btn_ChangeMinCutPartLength_Click(object sender, EventArgs e)
        {
            tsd.Drawing cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
            double minCutLength = new double();
            try
            {
                minCutLength = Convert.ToDouble(tb_MinimumCutPartLength.Text);
            }
            catch
            { }
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drobjenum.GetSize() == 0)
                ChangeDrawingMinCutPartLength(cur_dr, minCutLength);
            else
            {
                foreach (tsd.View view in drobjenum)
                {
                    view.Attributes.Shortening.MinimumLength = minCutLength;
                    view.Modify();
                }
            }
        }

        //Thực thi nút ChangeMinCutPartLength sau khi nhấn enter ở ô nhập cut part length
        private void tb_MinimumCutPartLength_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btn_ChangeMinCutPartLength_Click(this, new EventArgs());
        }
        public void ChangeDrawingMinCutPartLength(tsd.Drawing cur_dr, double minCutLength)
        {
            cur_dr = drawingHandler.GetActiveDrawing(); //Lay ban ve hien tai
            string Name = "TemporaryMacro.cs";//Temporary file name.
            string MacrosPath = string.Empty;//would store Path to script
            string script2 = "";
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
            if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
            if (minCutLength != 0)
            {
                if (cur_dr is tsd.AssemblyDrawing)
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""gr_adraw_on_off"", ""adraw_dial"");" +
                    @"akit.PushButton(""gr_adraw_view"", ""adraw_dial"");" +
                    @"akit.PushButton(""dv_on_off"", ""adv_dial"");" +
                    //@"akit.TabChange(""adv_dial"", ""contMain"", ""tabAttributes"");" +
                    //@"akit.ValueChange(""adv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                    @"akit.TabChange(""adv_dial"", ""contMain"", ""tabShortening"");" +
                    @"akit.ValueChange(""adv_dial"", ""gr_view_cut_on"", ""1"");" +         //chọn cho phép cut part
                    @"akit.ValueChange(""adv_dial"", ""gr_view_cut_skew_parts"", ""1"");" + //chọn cho phép cut part xiêng
                    @"akit.ValueChange(""adv_dial"",""gr_view_cut_min_dist"",""" + minCutLength + @""");" +
                    @"akit.PushButton(""dv_modify"", ""adv_dial"");" +
                    @" akit.PushButton(""gr_adraw_ok"", ""adraw_dial"");" +
                    "}}}";

                }
                else if (cur_dr is tsd.SinglePartDrawing)
                {
                    //Create a script as string
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                "{" +
                                " public class Script" +
                                    "{" +
                                        " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                        " {" +
                @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                @"akit.PushButton(""gr_wdraw_on_off"", ""wdraw_dial"");" +
                @"akit.PushButton(""gr_wdraw_view"", ""wdraw_dial"");" +
                @"akit.PushButton(""dv_on_off"", ""wdv_dial"");" +
                //@"akit.TabChange(""wdv_dial"", ""contMain"", ""tabAttributes"");" +
                //@"akit.ValueChange(""wdv_dial"", ""gr_dv_main_scale"", """ + (sender as Button).Text + @""");" +
                @"akit.TabChange(""wdv_dial"", ""contMain"", ""tabShortening"");" +
                @"akit.ValueChange(""wdv_dial"",""gr_view_cut_min_dist"",""" + minCutLength + @""");" +
                @"akit.PushButton(""dv_modify"", ""wdv_dial"");" +
                @" akit.PushButton(""gr_wdraw_ok"", ""wdraw_dial"");" +
                "}}}";
                }
            }
            //Save to file
            File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
            //Run it at Tekla Structures application
            Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
        }

        private void btn_ChangeSectionOrDetailName_Click(object sender, EventArgs e)
        {
            tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            foreach (tsd.DrawingObject item in drobjenum)
            {
                string Name = "TemporaryMacro.cs";//Temporary file name.
                string MacrosPath = string.Empty;//would store Path to script
                string script2 = "";
                string markText = tb_SectionDetailNamePrefix.Text + (sender as Button).Text;
                Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);//Find out where is actual macro directory.
                if (MacrosPath.IndexOf(';') > 0) { MacrosPath = MacrosPath.Remove(MacrosPath.IndexOf(';')); }
                if (item is tsd.SectionMarkBase)
                {
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""csym_on_off"", ""csym_dial"");" +
                    @"akit.ValueChange(""csym_dial"", ""csym_label"", """ + markText + @""");" +
                    @"akit.PushButton(""csym_modify"", ""csym_dial"");" +
                    @" akit.PushButton(""csym_cancel"", ""csym_dial"");" +
                    "}}}";
                    //Save to file
                    File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
                    //Run it at Tekla Structures application
                    Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
                }
                else if (item is tsd.DetailMark)
                {
                    script2 = "namespace Tekla.Technology.Akit.UserScript" +
                                    "{" +
                                    " public class Script" +
                                        "{" +
                                            " public static void Run(Tekla.Technology.Akit.IScript akit)" +
                                            " {" +
                    @"akit.Callback(""acmd_display_selected_drawing_object_dialog"", """", ""main_frame"");" +
                    @"akit.PushButton(""butDetailSymbol_on_off"", ""detail_dial"");" +
                    @"akit.ValueChange(""detail_dial"", ""lbltxtDetailLabelIndexStart"", """ + markText + @""");" +
                    @"akit.PushButton(""butDetailSymbol_modify"", ""detail_dial"");" +
                    @" akit.PushButton(""butDetailSymbol_cancel"", ""detail_dial"");" +
                    "}}}";
                    //Save to file
                    File.WriteAllText(Path.Combine(MacrosPath, Name), script2);
                    //Run it at Tekla Structures application
                    Tekla.Structures.Model.Operations.Operation.RunMacro("..\\" + Name);
                }
                else if (item is tsd.Plugin)
                {
                    //var myPugin = item as tsd.Plugin;
                    //myPugin.Select();
                    // myPugin.SetUserProperty("At_Revision", markText);
                    //myPugin.Modify();
                }
            }
        }

        #endregion


        private void timer1_Tick(object sender, EventArgs e)//kiểm tra có kết nối được với model không
        {
            try
            {
                model.GetInfo();//nếu model có mở thì sẽ lấy thông tin bản vẽ. nếu không sẽ tắt chương trình.
                GetDrawingInfor();//xem. Nếu bản vẽ mở thì load thông tin bản vẽ đang mở, nếu bv không mở thì không hiển thị thông tin nữa.
            }
            catch
            {
                this.Close();
            }
        }

        private void btn_CreateAngleByParts_Click(object sender, EventArgs e)
        {
            // Khai báo drawingHandler và model đã làm rồi (pulic static), nếu chưa có có thể thêm ở đây.
            //Khai báo workplanhandler
            try
            {
                Function function = new Function();
                tsm.WorkPlaneHandler wph = model.GetWorkPlaneHandler();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                CreateDimension createDimension = new CreateDimension(); //để tạo DIM : xem thêm class CreateDimension
                tsd.Drawing actived_dr = drawingHandler.GetActiveDrawing();
                tsd.DrawingObjectEnumerator drobjenum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                tsdui.Picker c_picker = drawingHandler.GetPicker();
                int dimSpace = Convert.ToInt32(tb_DimSpaceDim2Part.Text);
                int dimSpaceDim2Dim = Convert.ToInt32(tb_DimSpaceDim2Dim.Text);
                tsd.Part part1 = null;
            Begin:
                tsdui.Picker picker = drawingHandler.GetPicker();
                part1 = picker.PickObject("Select main part").Item1 as tsd.Part;
                tsm.Part mPart1 = model.SelectModelObject(part1.ModelIdentifier) as tsm.Part;
                //MessageBox.Show(mPart1.GetPartMark());
                tsd.ViewBase viewBase = part1.GetView();
                tsd.View view = viewBase as tsd.View;
                wph.SetCurrentTransformationPlane(new TransformationPlane(view.DisplayCoordinateSystem));
                t3d.Point pointA = mPart1.GetReferenceLine(true)[0] as t3d.Point;
                t3d.Point pointB = mPart1.GetReferenceLine(true)[1] as t3d.Point;

                foreach (tsd.DrawingObject drobj in drobjenum)
                {
                    if (drobj is tsd.Part)
                    {
                        //Convert về Part
                        tsd.Part drPart = drobj as tsd.Part;
                        //Lấy model part thông qua part model identifỉe
                        tsm.Part mPart2 = model.SelectModelObject(drPart.ModelIdentifier) as tsm.Part;
                        t3d.Point point1 = mPart2.GetReferenceLine(true)[0] as t3d.Point;
                        t3d.Point point2 = mPart2.GetReferenceLine(true)[1] as t3d.Point;

                        t3d.Point point1_2d = new t3d.Point(point1.X, point1.Y, 0);
                        t3d.Point point2_2d = new t3d.Point(point2.X, point2.Y, 0);
                        t3d.Point pointA_2d = new t3d.Point(point1.X, pointA.Y, 0);
                        t3d.Point pointB_2d = new t3d.Point(point2.X, pointB.Y, 0);

                        t3d.Line lineAB = new t3d.Line(pointA_2d, pointB_2d);
                        t3d.Line line12 = new t3d.Line(point1_2d, point2_2d);

                        LineSegment lineSegment = Intersection.LineToLine(lineAB, line12);
                        t3d.Point pointCenter = lineSegment.Point1;
                        //MessageBox.Show(Distance.PointToLine(point1, lineAB).ToString() + "_" + Distance.PointToLine(point2, lineAB).ToString());
                        if (Distance.PointToLine(point1_2d, lineAB) <= Distance.PointToLine(point2_2d, lineAB))
                            point1_2d = new t3d.Point(point2_2d.X, point2_2d.Y);
                        //function.InsertSymbol(viewBase, pointCenter);
                        //function.InsertLine(viewBase, point1, point2);
                        //MessageBox.Show(pointCenter.ToString() + "_" + point1.ToString() + "_" + pointA.ToString() + "_" + pointB.ToString());
                        AngleDimension myAngle1 = new AngleDimension(viewBase, pointCenter, point1_2d, pointA_2d, 200);
                        myAngle1.Attributes = new AngleDimensionAttributes(cb_DimentionAttributes.Text);
                        myAngle1.Attributes.Type = AngleTypes.AngleAtVertex;
                        myAngle1.Distance = 20;
                        myAngle1.Insert();
                        AngleDimension myAngle2 = new AngleDimension(viewBase, pointCenter, point1_2d, pointB_2d, 200);
                        myAngle2.Attributes = new AngleDimensionAttributes(cb_DimentionAttributes.Text);
                        myAngle2.Attributes.Type = AngleTypes.AngleAtVertex;
                        myAngle2.Distance = 20;
                        myAngle2.Insert();
                        if (Control.ModifierKeys == Keys.Shift)
                        {
                            if (myAngle1.GetAngle() < myAngle2.GetAngle())
                            {
                                myAngle1.Delete();
                            }
                            else myAngle2.Delete();
                        }
                        else
                        {
                            if (myAngle1.GetAngle() > myAngle2.GetAngle())
                            {
                                myAngle1.Delete();
                            }
                            else myAngle2.Delete();
                        }
                    }
                }
                drobjenum.Reset();
                wph.SetCurrentTransformationPlane(new TransformationPlane());
                actived_dr.CommitChanges();
                goto Begin;
            }
            catch
            { }
        }

        private void btn_DimFixedProfileGutter_Click(object sender, EventArgs e)
        {
            DimPickedPoints();
        }

        private void tb_DimSpaceDim2Part_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = (e.KeyChar == (char)Keys.Space);
            if (char.IsLetter(e.KeyChar)) //không cho phép nhập a-z
            {
                e.Handled = true;
            }
        }

        private void tb_DimSpaceDim2Part_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tb_DimSpaceDim2Part.Text))
                tb_DimSpaceDim2Part.Text = "0";
            else if (string.IsNullOrWhiteSpace(tb_DimSpaceDim2Dim.Text))
                tb_DimSpaceDim2Dim.Text = "0";
        }
    }
}


