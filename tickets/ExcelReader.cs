using System.Windows.Forms;
using System.Collections.Generic;
using SpreadsheetLight;
using System.Linq;
using System;

namespace tickets
{
    class ExcelReader
    {
        public static List<Unit> Read(string path)
        {
            var units = new List<Unit>();
            try {
                using (var doc = new SLDocument(path)) {
                    var stats = doc.GetWorksheetStatistics();

                    int ART = 1;
                    int NAME = 2;
                    int DIM = 0;
                    int NOTE = 4;
                    int QTY = 5;
                    int THICKNESS = 0;
                    int LENGHT = 0;
                    int WIDTH = 0;

                    int col = 0;
                    string header;
                    while ((header = doc.GetCellValueAsString(1, ++col).ToLower()) != "") {
                        switch (header) {
                            case "обозначение":
                            case "art":
                                ART = col;
                                break;
                            case "наименование":
                                NAME = col;
                                break;
                            case "размеры":
                                DIM = col;
                                break;
                            case "примечание":
                                NOTE = col;
                                break;
                            case "кол.":
                            case "кол":
                            case "к-во":
                                QTY = col;
                                break;
                            case "толщина":
                                THICKNESS = col;
                                break;
                            case "длина":
                                LENGHT = col;
                                break;
                            case "ширина":
                                WIDTH = col;
                                break;
                            default:
                                break;
                        }
                    }

                    string art, name, dim, th, len, wi, note;
                    int qty;
                    var stopStrings = new string[] { "", " ", "\t", "\n", "фурнитура", "метизы", "материал", "материалы", "лкм", "прочее", "заказные изделия" };

                    for (int row = stats.StartRowIndex + 2; row < stats.EndRowIndex; row++) {

                        art = doc.GetCellValueAsString(row, ART);
                        name = doc.GetCellValueAsString(row, NAME);
                        dim = DIM == 0 ? "" : doc.GetCellValueAsString(row, DIM);
                        len = LENGHT == 0 ? "" : doc.GetCellValueAsString(row, LENGHT);
                        wi = WIDTH == 0 ? "" : doc.GetCellValueAsString(row, WIDTH);
                        th = THICKNESS == 0 ? "" : doc.GetCellValueAsString(row, THICKNESS);
                        note = doc.GetCellValueAsString(row, NOTE);
                        qty = doc.GetCellValueAsInt32(row, QTY);

                        if (stopStrings.Contains(art.ToLower()))
                            break;

                        var unit = new Unit(art, name, dim, th, len, wi, note, qty);
                        units.Add(unit);
                    }
                    doc.CloseWithoutSaving();
                }
            }
            catch (System.IO.IOException e) {
                Console.WriteLine(e.Message);
                return null;
            }

            Console.WriteLine(string.Format("Из файла спецификации \"{0}\"", path));
            Console.WriteLine("загружено: ");
            foreach (var unit in units) {
                Console.WriteLine(unit.ToString());
            }

            return units;
        }
    }
}