using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KmsReportClient.DgvHeaderGenerator
{
   public class Header
    {
        public List<Header> Children { get; set; }

        public string Name { get; set; }

        public int X { get; set; }

        public int Y { get; set; }

        public int Width { get; set; }

        public int Height { get; set; }

        public int ColumnId { get; set; }

        public Header()
        {
            Name = string.Empty;
            Children = new List<Header>();
            ColumnId = -1;
        }

        public void Measure(DataGridView objGrid, int iY, int iHeight)
        {
            Width = 0;
            if (Children.Count > 0)
            {
                int tempY = string.IsNullOrEmpty(Name.Trim()) ? iY : iY + iHeight;
                bool columnWidthSet = false;
                foreach (Header child in Children)
                {
                    child.Measure(objGrid, tempY, iHeight);
                    Width += child.Width;
                    if (!columnWidthSet && Width > 0)
                    {
                        ColumnId = child.ColumnId;
                        columnWidthSet = true;
                    }
                }
            }
            else if (-1 != ColumnId && objGrid.Columns[ColumnId].Visible)
            {
                Width = objGrid.Columns[ColumnId].Width;
            }
            Y = iY;
            if (Children.Count == 0)
            {
                Height = objGrid.ColumnHeadersHeight - iY;
            }
            else
            {
                Height = iHeight;
            }
        }

        //public void AcceptRenderer(StackedHeaderDecorator objRenderer, DataGridView objGrid, int iY)
        //{
        //    foreach (Header children in Children)
        //    {
        //        children.AcceptRenderer(objRenderer, objGrid, iY);
        //    }
        //    if (-1 != ColumnId && !string.IsNullOrEmpty(Name.Trim()))
        //    {
        //        objRenderer.Render(this);
        //    }

        //}

        public void AcceptRenderer(StackedHeaderDecorator objRenderer)
        {
            foreach (Header objChild in Children)
            {
                objChild.AcceptRenderer(objRenderer);
            }
            if (-1 != ColumnId && !string.IsNullOrEmpty(Name.Trim()))
            {
                objRenderer.Render(this);
            }

        }
    }
}
