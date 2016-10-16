using System;
using SpreadsheetLight;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

namespace tickets
{
    class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            List<Unit> units;
            using (var dlg = new OpenFileDialog()) {

                dlg.Title = "Выбор файла спецификации";
                var dr = dlg.ShowDialog();

                if (dr != DialogResult.OK)
                    return;

                units = ExcelReader.Read(dlg.FileName);
                if (units != null) {
                    makeStickers(units, dlg.FileName);
                }
                Console.WriteLine("Для завершения нажмите любую клавишу.");
                Console.ReadKey();
            }
        }

        static void makeStickers(List<Unit> units, string spName)
        {
            using (SLDocument doc = new SLDocument()) {
                var style_12 = new SLStyle();
                style_12.SetFont("arial", 12);
                style_12.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                style_12.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);

                SLPageSettings ps = new SLPageSettings();
                ps.TopMargin = 0;
                ps.BottomMargin = 0;
                ps.LeftMargin = 0;
                ps.RightMargin = 0;

                doc.SetPageSettings(ps);

                int col = 1; int row = 1;
                foreach (Unit unit in units) {
                    for (int qty = 0; qty < unit.qty; qty++) {
                        if (col == 5) {
                            col = 1;
                            row++;
                        }
                        doc.SetColumnWidth(col, COLUMN_WIDTH_FOR_ARIAL_14);
                        doc.SetRowHeight(row, ROW_HEIGHT_FOR_ARIAL_14);
                        doc.SetCellStyle(row, col, style_12);
                        doc.SetCellValue(row, col++, unit.GetSticker());
                    }
                }

                using (var dlg = new SaveFileDialog()) {
                    dlg.Title = "Сохранение файла наклеек";
                    dlg.FileName = "Наклейки -- " + spName.Split('\\').Last();
                    var res = dlg.ShowDialog();
                    if (res != DialogResult.OK)
                        return;
                    doc.SaveAs(dlg.FileName);
                    Console.WriteLine("Сохранено в " + dlg.FileName);
                }
            }
        }

        static double ROW_HEIGHT_FOR_ARIAL_14 = 84;
        static double COLUMN_WIDTH_FOR_ARIAL_14 = 27.25;
    }

}