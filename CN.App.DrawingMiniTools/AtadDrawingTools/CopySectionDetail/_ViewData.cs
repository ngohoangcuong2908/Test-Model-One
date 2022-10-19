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
    class _ViewData
    {
        public TSD.View view;
        public T3D.Point offsetPoint;

        public List<TSD.SectionMark> sectionMarks;
        public List<TSD.DetailMark> detailMarks;

        public _ViewData()
        {
            sectionMarks = new List<TSD.SectionMark>();
            detailMarks = new List<TSD.DetailMark>();
        }

        public void setView(TSD.View currentView, T3D.Point point)
        {
            view = currentView;
            offsetPoint = point;
        }

        public void addOneObject(TSD.DrawingObject dro)
        {
            if (dro is TSD.SectionMark)
            {
                TSD.SectionMark current = dro as TSD.SectionMark;
                sectionMarks.Add(current);
            }

            else if (dro is TSD.DetailMark)
            {
                TSD.DetailMark current = dro as TSD.DetailMark;
                detailMarks.Add(current);
            }
        }

        public string countObjects()
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine(" ");
            message.AppendLine("Section marks: " + sectionMarks.Count);
            message.AppendLine("Detail marsk: " + detailMarks.Count);
            return message.ToString();
        }
    }
}
