using System;
using System.Collections.Generic;
using System.Drawing;
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

        public void Measure(DataGridView objGrid, int startY, int levelHeight, int totalHeaderHeight)
        {
            Width = 0;
            Y = startY;

            if (Children.Count > 0)
            {
                // Если у узла есть имя — резервируем место под него
                int childrenStartY = string.IsNullOrWhiteSpace(Name) ? startY : startY + levelHeight;

                // Рекурсивно измеряем детей
                foreach (Header child in Children)
                {
                    child.Measure(objGrid, childrenStartY, levelHeight, totalHeaderHeight);
                    Width += child.Width;

                    // Устанавливаем ColumnId в первый ВИДИМЫЙ дочерний столбец
                    if (ColumnId == -1 && child.ColumnId != -1 && child.Width > 0)
                    {
                        ColumnId = child.ColumnId;
                    }
                }

                Height = string.IsNullOrWhiteSpace(Name) ? (totalHeaderHeight - startY) : levelHeight;
            }
            else if (ColumnId != -1 && ColumnId < objGrid.Columns.Count && objGrid.Columns[ColumnId].Visible)
            {
                // Листовой узел: ширина = ширина столбца, высота = остаток до низа
                Width = objGrid.Columns[ColumnId].Width;
                Height = totalHeaderHeight - startY;
            }
            else
            {
                // Скрытый или невалидный столбец
                Width = 0;
                Height = 0;
            }
        }

        //public void Measure(DataGridView objGrid, int startY, int levelHeight, int totalHeaderHeight, int level = 0)
        //{
        //    Width = 0;
        //    Y = startY;

        //    // 🔹 Кастомная высота: для level == 1 (второй уровень — описание) — увеличиваем
        //    int actualLevelHeight = levelHeight;
        //    if (level == 1 && objGrid.Parent?.Name == "Page140n")
        //    {
        //        actualLevelHeight = levelHeight * 3; // или 2, 2.5 — подберите
        //    }

        //    if (Children.Count > 0)
        //    {
        //        int childrenStartY = string.IsNullOrWhiteSpace(Name) ? startY : startY + actualLevelHeight;

        //        foreach (Header child in Children)
        //        {
        //            child.Measure(objGrid, childrenStartY, levelHeight, totalHeaderHeight, level + 1);
        //            Width += child.Width;

        //            if (ColumnId == -1 && child.ColumnId != -1 && child.Width > 0)
        //                ColumnId = child.ColumnId;
        //        }

        //        Height = string.IsNullOrWhiteSpace(Name) ? (totalHeaderHeight - startY) : actualLevelHeight;
        //    }
        //    else if (ColumnId >= 0 && ColumnId < objGrid.Columns.Count && objGrid.Columns[ColumnId].Visible)
        //    {
        //        Width = objGrid.Columns[ColumnId].Width;

        //        // Измеряем высоту текста с переносом
        //        using (var g = objGrid.CreateGraphics())
        //        {
        //            var font = objGrid.ColumnHeadersDefaultCellStyle.Font;
        //            var sf = new StringFormat
        //            {
        //                Alignment = StringAlignment.Center,
        //                LineAlignment = StringAlignment.Near,
        //                Trimming = StringTrimming.Word,
        //                FormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip
        //            };

        //            SizeF size = g.MeasureString(Name, font, Width, sf);
        //            Height = Math.Min((int)Math.Ceiling(size.Height), totalHeaderHeight - startY);
        //        }
        //    }
        //    else
        //    {
        //        Width = 0;
        //        Height = 0;
        //    }
        //}

        public void AcceptRenderer(StackedHeaderDecorator renderer)
        {
            // Сначала отрисовываем текущий узел (верхние уровни)
            if (ColumnId != -1 && !string.IsNullOrWhiteSpace(Name) && Width > 0 && Height > 0)
            {
                renderer.Render(this);
            }

            // Затем рекурсивно обрабатываем детей (нижние уровни поверх)
            foreach (Header child in Children)
            {
                child.AcceptRenderer(renderer);
            }
        }
    }
}
