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
