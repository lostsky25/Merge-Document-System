using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace MDS
{
    class Marge
    {
        private Excel.Application applicationExcel;
        private Excel.Workbook workbookExcel;
        private Excel.Worksheet sheetExcel;
        private string filePathWord;
        private string filePathExcel;
        private readonly String pathWordTemplate = "teamplate_word.docx";

        private List<string> listNumberOfPlace;
        private List<string> listFIO;
        private List<string> listAddress;
        private List<string> listCost;
        private List<string> listPeriod;
        private List<string> listTypeOfPayment;
        private List<string> listEmail;

        private int countClients;

        public Marge(string filePathExcel)
        {
            this.filePathWord = pathWordTemplate;
            this.filePathExcel = filePathExcel;

            applicationExcel = new Excel.Application();

            workbookExcel = applicationExcel.Workbooks.Open(filePathExcel);
            sheetExcel = workbookExcel.ActiveSheet;
        }

        public void Clear()
        {
            listNumberOfPlace.Clear();
            listFIO.Clear();
            listAddress.Clear();
            listCost.Clear();
            listPeriod.Clear();
            listTypeOfPayment.Clear();
            listEmail.Clear();
            countClients = 0;
        }

        public List<string> getEmails()
        {
            return listEmail;
        }

        public List<string> getFIO()
        {
            return listFIO;
        }
        private bool CheckColumn(string nameColumn)
        {
            Excel.Range searchedRange = applicationExcel.get_Range("A1", "XFD1048576");
            Excel.Range currentFind = searchedRange.Find(nameColumn);

            if (currentFind != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private List<string> GetAllColumn(string nameColumn)
        {
            Excel.Range range, r2;
            range = sheetExcel.UsedRange.Find(nameColumn);
            List<string> cellsData = new List<string>();
            r2 = sheetExcel.Cells;
            int n_c = range.Column;
            int n_r = range.Row;

            int last = sheetExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            range = sheetExcel.get_Range("B1:B" + last);

            int cnt = 0;
            foreach (Excel.Range element in range.Cells)

            {
                if (element.Value2 != null)
                {
                    cnt = cnt + 1;
                }

            }

            for (int i = 1; i < cnt; i++)
            {
                cellsData.Add(Convert.ToString(((Excel.Range)r2[n_r + i, n_c]).Value));
            }

            return cellsData;
        }

        public bool InitializeClients()
        {
            bool validSheet = true;

            foreach (string nameColumn in new List<string>() { "№ машиноместа", "ФИО", "Адрес машиноместа", "Сумма", "Период", "Назначение платежа", "e-mail" })
            {
                if (!CheckColumn(nameColumn))
                {
                    validSheet = false;
                }
            }

            if (validSheet)
            {
                listNumberOfPlace = GetAllColumn("№ машиноместа");
                listFIO = GetAllColumn("ФИО");
                listAddress = GetAllColumn("Адрес машиноместа");
                listCost = GetAllColumn("Сумма");
                listPeriod = GetAllColumn("Период");
                listTypeOfPayment = GetAllColumn("Назначение платежа");
                listEmail = GetAllColumn("e-mail");
                countClients = listFIO.Count;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheetExcel);
            workbookExcel.Close(true);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbookExcel);
            applicationExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(applicationExcel);

            return validSheet;
        }
        public static void SaveLog(string templatePath, string outPath, List<string> names, List<string> senderEmails, List<string> clientEmails, List<DateTime> sendTime, List<string> status)
        {
            Excel.Application applicationExcel = new Excel.Application();
            Excel.Workbook workbookExcel = applicationExcel.Workbooks.Open(Directory.GetCurrentDirectory() + "\\" + templatePath);
            Excel.Worksheet sheetExcel = workbookExcel.ActiveSheet;

            if (applicationExcel != null)
            {
                for (int i = 0; i < names.Count; i++)
                {
                    sheetExcel.Cells[i + 2, 1] = names[i];
                    sheetExcel.Cells[i + 2, 2] = senderEmails[i];
                    sheetExcel.Cells[i + 2, 3] = clientEmails[i];
                    sheetExcel.Cells[i + 2, 4] = sendTime[i].ToString("dd.MM.yyyy");
                    sheetExcel.Cells[i + 2, 5] = sendTime[i].ToString("HH:mm:ss");
                    sheetExcel.Cells[i + 2, 6] = status[i];

                    if (status[i].Contains("Успешно"))
                    {
                        sheetExcel.Cells[i + 2, 6].Interior.Color = Excel.XlRgbColor.rgbLightGreen;
                    }
                    else
                    {
                        sheetExcel.Cells[i + 2, 6].Interior.Color = Excel.XlRgbColor.rgbDarkRed;
                    }

                }

                workbookExcel.SaveAs(outPath + @"\" + DateTime.Now.ToString("Отчет_dd_MM_yyyy_HH_mm_ss").ToString() + ".xlsx",
                    Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Missing.Value, Missing.Value,
                    false,
                    false,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution,
                    true,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value);

                workbookExcel.Close();
                applicationExcel.Quit();
            }


            Marshal.ReleaseComObject(applicationExcel);
            Marshal.ReleaseComObject(workbookExcel);
            Marshal.ReleaseComObject(sheetExcel);
        }
        public int getCount()
        {
            return countClients;
        }

        string getMonthName(int month)
        {
            string monthName;
            switch (month)
            {
                case 1:
                    monthName = "Янв";
                    break;
                case 2:
                    monthName = "Фев";
                    break;
                case 3:
                    monthName = "Мар";
                    break;
                case 4:
                    monthName = "Апр";
                    break;
                case 5:
                    monthName = "Май";
                    break;

                case 6:
                    monthName = "Июнь";
                    break;

                case 7:
                    monthName = "Июль";
                    break;

                case 8:
                    monthName = "Авг";
                    break;

                case 9:
                    monthName = "Сен";
                    break;

                case 10:
                    monthName = "Окт";
                    break;

                case 11:
                    monthName = "Ноя";
                    break;

                case 12:
                    monthName = "Дек";
                    break;
                default:
                    monthName = "";
                    break;
            }

            return monthName;
        }

        bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false; // suggested by @TK-421
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }

        //Creeate the Doc Method
        public /*Tuple<List<string>, List<string>>*/List<string> getMargedEW(string outPath)
        {
            List<string> filesPath = new List<string>();

            if (Directory.Exists((string)outPath) && File.Exists((string)filePathExcel))
            {
                for (int i = 0; i < listFIO.Count; i++)
                {
                    string[] fieldName = { "$FIO", "$NOP", "$COST", "$ADDR", "$DAY", "$MONTH", "$YEAR" };
                    var parsedDate = DateTime.Parse(listPeriod.ElementAt(i));

                    string _year_short = parsedDate.Year.ToString().Remove(0, 2).PadLeft(2, '0');
                    string _month = getMonthName(parsedDate.Month).PadLeft(2, '0');
                    string _day = parsedDate.Day.ToString().PadLeft(2, '0');



                    string[] fieldValue = {
                            listFIO.ElementAt(i),
                            listNumberOfPlace.ElementAt(i),
                            listCost.ElementAt(i),
                            listAddress.ElementAt(i),
                            _day,
                            _month,
                            _year_short,
                        };



                    Spire.Doc.Document documentWord = new Document(filePathWord);

                    documentWord.Replace(fieldName[0], fieldValue[0], false, true);
                    documentWord.Replace(fieldName[1], fieldValue[1], false, true);
                    documentWord.Replace(fieldName[2], fieldValue[2], false, true);
                    documentWord.Replace(fieldName[3], fieldValue[3], false, true);
                    documentWord.Replace(fieldName[4], fieldValue[4], false, true);
                    documentWord.Replace(fieldName[5], fieldValue[5], false, true);
                    documentWord.Replace(fieldName[6], fieldValue[6], false, true);

                    filesPath.Add(listFIO.ElementAt(i).Replace(" ", "_") + "_" + i.ToString() + @".docx");
                    documentWord.SaveToFile(outPath + "\\" + listFIO.ElementAt(i).Replace(" ", "_") + "_" + i.ToString() + @".docx", FileFormat.Docx2010);

                    documentWord.Close();
                }
            }
            countClients = listFIO.Count;
            //Clear();
            return filesPath;
        }
    }
}
