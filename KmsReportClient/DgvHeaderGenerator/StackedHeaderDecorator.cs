using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KmsReportClient.DgvHeaderGenerator
{
    public class StackedHeaderDecorator
    {

        private readonly IStackedHeaderGenerator objStackedHeaderGenerator = StackedHeaderGenerator.Instance;
        private Graphics objGraphics;
        private readonly DataGridView objDataGrid;
        private Header objHeaderTree;
        private int iNoOfLevels;
        private readonly StringFormat objFormat;

        public StackedHeaderDecorator(DataGridView objDataGrid)
        {
            this.objDataGrid = objDataGrid;
            objFormat = new StringFormat();
            objFormat.Alignment = StringAlignment.Center;
            objFormat.LineAlignment = StringAlignment.Far;

            Type dgvType = objDataGrid.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(objDataGrid, true, null);

            objDataGrid.Scroll += (objDataGrid_Scroll);
            objDataGrid.Paint += objDataGrid_Paint;
            objDataGrid.ColumnRemoved += objDataGrid_ColumnRemoved;
            objDataGrid.ColumnAdded += objDataGrid_ColumnAdded;
            objDataGrid.ColumnWidthChanged += objDataGrid_ColumnWidthChanged;
            objHeaderTree = objStackedHeaderGenerator.GenerateStackedHeader(objDataGrid);

        }

        public StackedHeaderDecorator(IStackedHeaderGenerator objStackedHeaderGenerator, DataGridView objDataGrid)
           : this(objDataGrid)
        {
            this.objStackedHeaderGenerator = objStackedHeaderGenerator;
        }


        void objDataGrid_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Refresh();
        }

        void objDataGrid_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {

            RegenerateHeaders();
            Refresh();
        }

        void objDataGrid_ColumnRemoved(object sender, DataGridViewColumnEventArgs e)
        {
            RegenerateHeaders();
            Refresh();
        }

        void objDataGrid_Paint(object sender, PaintEventArgs e)
        {
            iNoOfLevels = NoOfLevels(objHeaderTree);
            objGraphics = e.Graphics;
            objDataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            if (objDataGrid.Parent.Name == "PageCadre")
            {
                objDataGrid.ColumnHeadersHeight = 210;
                //objDataGrid.DefaultCellStyle.BackColor = Color.FromArgb(253, 233, 217);
            }
            if (objDataGrid.Parent.Name == "PageMonitoringVCR")
            {
                objDataGrid.ColumnHeadersHeight += 50;
            }
            if (objDataGrid.Parent.Name == "PageT5Newborn")
            {
                objDataGrid.ColumnHeadersHeight = 150; // или 100 — подберите визуально
            }
            //else if (objDataGrid.Parent.Name == "PageQuantity")
            //{
            //    objDataGrid.ColumnHeadersHeight = 163;
            //}
            //else
            //{
            //objDataGrid.ColumnHeadersHeight = objDataGrid.ColumnHeadersHeight;
            //}

            if (null != objHeaderTree)
            {
                RenderColumnHeaders();
            }
        }

        void objDataGrid_Scroll(object sender, ScrollEventArgs e)
        {
            Refresh();
        }

        private void Refresh()
        {
            Rectangle rtHeader = objDataGrid.DisplayRectangle;
            objDataGrid.Invalidate(rtHeader);
        }

        private void RegenerateHeaders()
        {
            objHeaderTree = objStackedHeaderGenerator.GenerateStackedHeader(objDataGrid);
        }

        private void RenderColumnHeaders()
        {
            // === ВАЛИДАЦИЯ ===
            if (objDataGrid == null || objHeaderTree?.Children == null || iNoOfLevels <= 0)
            {
                Console.WriteLine("RenderColumnHeaders: Invalid state (null grid, headers or levels <= 0)");
                return;
            }

            if (objDataGrid.ColumnHeadersHeight <= 0)
            {
                Console.WriteLine("RenderColumnHeaders: ColumnHeadersHeight is zero or negative");
                return;
            }

            // === ОЧИСТКА ФОНА ===
            var headerArea = new Rectangle(
                objDataGrid.DisplayRectangle.X,
                objDataGrid.DisplayRectangle.Y,
                objDataGrid.DisplayRectangle.Width,
                objDataGrid.ColumnHeadersHeight
            );

            using (var bgBrush = new SolidBrush(objDataGrid.ColumnHeadersDefaultCellStyle.BackColor))
            {
                objGraphics.FillRectangle(bgBrush, headerArea);
            }

            // === РАСЧЁТ ГЕОМЕТРИИ ===
            const int TOP_PADDING = 3; // Настроено под визуальную гармонию (было 5)
            int baseLevelHeight = objDataGrid.ColumnHeadersHeight / iNoOfLevels;
            int remainder = objDataGrid.ColumnHeadersHeight % iNoOfLevels; // Остаток для последнего уровня

            // === ОБХОД КОРНЕВЫХ УЗЛОВ ===
            foreach (Header child in objHeaderTree.Children)
            {
                try
                {
                    // Передаём ОБЩУЮ высоту для корректного расчёта остатка в листовых узлах
                    child.Measure(
                        objDataGrid,
                        TOP_PADDING,
                        baseLevelHeight,
                        objDataGrid.ColumnHeadersHeight
                    );

                    // Префиксный обход гарантирует: верхние уровни рисуются ПЕРВЫМИ (под дочерними)
                    child.AcceptRenderer(this);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error rendering header '{child?.Name}': {ex.Message}");
                    // Продолжаем обработку остальных заголовков
                }
            }

            // === ДОП: Рисуем разделитель под заголовками (опционально) ===
            using (var linePen = new Pen(objDataGrid.GridColor, 1))
            {
                int bottomY = objDataGrid.DisplayRectangle.Y + objDataGrid.ColumnHeadersHeight - 1;
                objGraphics.DrawLine(
                    linePen,
                    objDataGrid.DisplayRectangle.X,
                    bottomY,
                    objDataGrid.DisplayRectangle.Right,
                    bottomY
                );
            }
        }

        public void Render(Header objHeader)
        {
            if (objHeader.Children.Count == 0)
            {
                Rectangle r1 = objDataGrid.GetColumnDisplayRectangle(objHeader.ColumnId, true);
                if (r1.Width == 0)
                {
                    return;
                }
                r1.Y = objHeader.Y;
                r1.Width += 1;
                r1.X -= 1;
                r1.Height = objHeader.Height;
                objGraphics.SetClip(r1);

                if (r1.X + objDataGrid.Columns[objHeader.ColumnId].Width < objDataGrid.DisplayRectangle.Width)
                {
                    r1.X -= (objDataGrid.Columns[objHeader.ColumnId].Width - r1.Width);
                }
                r1.X -= 1;
                r1.Width = objDataGrid.Columns[objHeader.ColumnId].Width;
                objGraphics.DrawRectangle(Pens.Gray, r1);
                objGraphics.DrawString(objHeader.Name,
                                       objDataGrid.ColumnHeadersDefaultCellStyle.Font,
                                       new SolidBrush(objDataGrid.ColumnHeadersDefaultCellStyle.ForeColor),
                                       r1,
                                       objFormat);
                objGraphics.ResetClip();
            }
            else
            {
                int x = objDataGrid.RowHeadersWidth;
                for (int i = 0; i < objHeader.Children[0].ColumnId; ++i)
                {
                    if (objDataGrid.Columns[i].Visible)
                    {
                        x += objDataGrid.Columns[i].Width;
                    }
                }
                if (x > (objDataGrid.HorizontalScrollingOffset + objDataGrid.DisplayRectangle.Width - 5))
                {
                    return;
                }

                Rectangle r1 = objDataGrid.GetCellDisplayRectangle(objHeader.Children[0].ColumnId, -1, true);
                //Rectangle r1 = objDataGrid.GetCellDisplayRectangle(objHeader.ColumnId, -1, true);
                r1.Y = objHeader.Y;
                r1.Height = objHeader.Height;
                r1.Width = objHeader.Width + 1;
                if (r1.X < objDataGrid.RowHeadersWidth)
                {
                    r1.X = objDataGrid.RowHeadersWidth;
                }
                r1.X -= 1;
                objGraphics.SetClip(r1);
                r1.X = x - objDataGrid.HorizontalScrollingOffset;
                r1.Width -= 1;
                objGraphics.DrawRectangle(Pens.Gray, r1);
                r1.X -= 1;
                objGraphics.DrawString(objHeader.Name, objDataGrid.ColumnHeadersDefaultCellStyle.Font,
                                       new SolidBrush(objDataGrid.ColumnHeadersDefaultCellStyle.ForeColor),
                                       r1, objFormat);
                objGraphics.ResetClip();
            }

        }

        private int NoOfLevels(Header header)
        {
            int level = 0;
            foreach (Header child in header.Children)
            {
                int temp = NoOfLevels(child);
                level = temp > level ? temp : level;
            }
            return level + 1;
        }
    }
}
