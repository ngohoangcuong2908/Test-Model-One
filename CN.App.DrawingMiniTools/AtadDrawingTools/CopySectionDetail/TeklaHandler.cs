using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Tekla.Structures;
using TSD = Tekla.Structures.Drawing;
using TSM = Tekla.Structures.Model;
using T3D = Tekla.Structures.Geometry3d;

namespace CN.App.DrawingTools.CopySectionDetail
{
    class TeklaHandler
    {
        public static _ViewData getInputData()
        {
            _ViewData view_data = new _ViewData();
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();

            if (drawingHandler.GetConnectionStatus())
            {
                TSD.ContainerView sheet = drawingHandler.GetActiveDrawing().GetSheet();

                TSD.UI.Picker picker = drawingHandler.GetPicker();
                T3D.Point viewPoint = null;
                TSD.ViewBase curView = null;


                picker.PickPoint("Select point in drawing view", out viewPoint, out curView);

                if (curView is TSD.View)
                {
                    view_data.setView(curView as TSD.View, viewPoint);
                }
                else
                {
                    throw new OperationCanceledException();
                }

                TSD.DrawingObjectEnumerator selectedObjects = drawingHandler.GetDrawingObjectSelector().GetSelected();
                populate(view_data, selectedObjects);
            }

            return view_data;
        }

        public static _ViewData getOutputData()
        {
            _ViewData view_data = new _ViewData();
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();

            if (drawingHandler.GetConnectionStatus())
            {
                TSD.ContainerView sheet = drawingHandler.GetActiveDrawing().GetSheet();

                TSD.UI.Picker picker = drawingHandler.GetPicker();
                T3D.Point viewPoint = null;
                TSD.ViewBase curView = null;

                picker.PickPoint("Pick one point", out viewPoint, out curView);

                if (curView is TSD.View)
                {
                    view_data.setView(curView as TSD.View, viewPoint);
                }
                else
                {
                    throw new OperationCanceledException();
                }
            }


            return view_data;

        }

        public static void populate(_ViewData data, TSD.DrawingObjectEnumerator all)
        {
            int i = 0;
            int tot = all.GetSize();

            foreach (TSD.DrawingObject one in all)
            {
                i++;

                if (one is TSD.SectionMark || one is TSD.DetailMark)
                {
                    data.addOneObject(one);
                }
            }
        }

        public static void copyView(_ViewData input, _ViewData output)
        {
            TSD.DrawingHandler drawingHandler = new TSD.DrawingHandler();

            if (drawingHandler.GetConnectionStatus())
            {
                createSectionMarks(input, output);
                createDetailMarks(input, output);
            }
        }

        private static void createSectionMarks(_ViewData input, _ViewData output)
        {
            List<TSD.SectionMark> inputSections = input.sectionMarks;

            foreach (TSD.SectionMark inputSection in inputSections)
            {
                T3D.Point leftPoint = applyOffset(input, output, inputSection.LeftPoint);
                T3D.Point rightPoint = applyOffset(input, output, inputSection.RightPoint);

                TSD.SectionMark outputSectionMark = new TSD.SectionMark(output.view, leftPoint, rightPoint, inputSection.Attributes);
                outputSectionMark.Insert();
            }
        }

        private static void createDetailMarks(_ViewData input, _ViewData output)
        {
            List<TSD.DetailMark> inputDetails = input.detailMarks;

            foreach (TSD.DetailMark inputDetail in inputDetails)
            {
                T3D.Point centerPoint = applyOffset(input, output, inputDetail.CenterPoint);
                T3D.Point boundaryPoint = applyOffset(input, output, inputDetail.BoundaryPoint);
                T3D.Point labelPoint = applyOffset(input, output, inputDetail.LabelPoint);

                TSD.DetailMark outputDetailMark = new TSD.DetailMark(output.view, centerPoint, boundaryPoint, labelPoint, inputDetail.Attributes);
                outputDetailMark.Insert();
            }
        }

        public static T3D.Point applyOffset(_ViewData input, _ViewData output, T3D.Point point)
        {
            double X = point.X - input.offsetPoint.X + output.offsetPoint.X;
            double Y = point.Y - input.offsetPoint.Y + output.offsetPoint.Y;

            T3D.Point tr = new T3D.Point(X, Y, output.offsetPoint.Z);
            return tr;
        }
    }
}
