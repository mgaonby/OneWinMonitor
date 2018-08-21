using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace OneWinMonitor
{
    public partial class showResultsForm : Form
    {
        DateTime beginDate, endDate;
        List<DateTime> data;
        int[] daysAverage;
        int[,] daysCount;
        string connectionString;
        string procedureName;
        string thisArea;
        public showResultsForm()
        {
            InitializeComponent();
        }
        public showResultsForm(DateTime beginDate, DateTime endDate, string area, string fullAreaName, string procedureName = "")
        {
            InitializeComponent();
                connectionString = String.Format(@"Server=tcp:172.16.209.208, 1433; Database=C:\OneWin\{0}\app_data\AISMINSKBASE.mdf;User Id = sa; Password = jlyjjryj; Integrated Security=false", area);
            this.beginDate = beginDate;
            this.endDate = endDate;
            this.procedureName = procedureName;
            data = new List<DateTime>();
            daysAverage = new int[7];
            daysCount = new int[18, 6];
            thisArea = area;
            for (int i = 0; i < daysCount.GetLength(0); i++)
            {
                for (int j = 0; j < daysCount.GetLength(1); j++)
                {
                    daysCount[i, j] = 0;
                }
            }
            for (int i = 0; i < daysAverage.Length; i++)
            {
                daysAverage[i] = 0;
            }
            selectedPeriod.Text = String.Format("{0} - {1}, {2} район", beginDate.ToShortDateString(), endDate.ToShortDateString(), fullAreaName);
            this.Text = String.Format("{0} - {1}, {2} район", beginDate.ToShortDateString(), endDate.ToShortDateString(), fullAreaName);

        }
        private void getStatistics()
        {
            for (int i = 0; i < resultDataGrid.ColumnCount; i++)
            {
                string s = resultDataGrid.Columns[i].HeaderText.ToString();
                for (int j = 0; j < resultDataGrid.RowCount-1; j++)
                {

                    
                    s = s.Remove(s.IndexOf(','));
                    DateTime newdate = DateTime.Parse(s);
                        daysCount[(int)newdate.DayOfWeek-1, i]+= (int)resultDataGrid.Rows[j].Cells[i].Value;
                }
            }
        }
        private void showResultsForm_Load(object sender, EventArgs e)
        {
            this.resultDataGrid.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 14);
            this.resultDataGrid.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 14);
            resultDataGrid.RowHeadersDefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 12);
            
            selectRowsFromDB();
            drawTable();
            loadDataInTable();

        }
        void loadDataInTable()
        {
            for (int i = 0; i < resultDataGrid.Columns.Count; i++)
            {
                for (int j = 0; j < resultDataGrid.Rows.Count; j++)
                {
                    resultDataGrid.Rows[j].Cells[i].Value = 0;
                }
            }
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < resultDataGrid.Columns.Count; j++)
                {
                    if (data[i].ToShortDateString() == resultDataGrid.Columns[j].Name)
                    {
                        int temp = (int)resultDataGrid[j, 17].Value;
                        temp++;
                        resultDataGrid[j, 17].Value = temp;
   
                    }
                }
            }
            for (int i = 0; i < resultDataGrid.Columns.Count; i++)
            {
                if ((int)resultDataGrid[i, 17].Value == 0)
                {
                    resultDataGrid.Columns.Remove(resultDataGrid.Columns[i].Name);
                    i--;
                }
            }
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < resultDataGrid.Columns.Count; j++)
                {
                    if (data[i].ToShortDateString() == resultDataGrid.Columns[j].Name)
                    {
                        int temp = (int)resultDataGrid[j, data[i].Hour - 5].Value;
                        resultDataGrid[j, data[i].Hour - 5].Value = ++temp;
                        int rowIndex = data[i].Hour - 5;
                        try
                        {
                            daysCount[ rowIndex, (int)data[i].DayOfWeek-1] ++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(rowIndex.ToString() + " " + i);
                        }
                    }
                }
            }
            for (int i = 0; i < resultDataGrid.Rows.Count; i++)
            {
                bool isEmptyRow = true;
                for (int j = 0; j < resultDataGrid.Columns.Count; j++)
                {
                    if ((int)resultDataGrid[j, i].Value != 0)
                    {
                        isEmptyRow = false;
                        break;
                    }
                }
                if (isEmptyRow)
                {
                    resultDataGrid.Rows.RemoveAt(i);
                    i--;
                }
            }
        }
        void drawTable()
        {
            TimeSpan diff = endDate - beginDate;
            DateTime tempBeginDate = beginDate;
            for (int i = 0; i < diff.Days; i++)
            {
                resultDataGrid.Columns.Add(tempBeginDate.ToShortDateString(), tempBeginDate.ToShortDateString() + ", "
                    + CultureInfo.GetCultureInfo("ru-RU").DateTimeFormat.GetDayName(tempBeginDate.DayOfWeek));
                resultDataGrid[i, 0].Value = 0;
                resultDataGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                tempBeginDate = tempBeginDate.AddDays(1);
            }
            resultDataGrid.Rows.Add(17);
            resultDataGrid.Rows[0].HeaderCell.Value = "05:00 - 06:00";
            resultDataGrid.Rows[1].HeaderCell.Value = "06:00 - 07:00";
            resultDataGrid.Rows[2].HeaderCell.Value = "07:00 - 08:00";
            resultDataGrid.Rows[3].HeaderCell.Value = "08:00 - 09:00";
            resultDataGrid.Rows[4].HeaderCell.Value = "09:00 - 10:00";
            resultDataGrid.Rows[5].HeaderCell.Value = "10:00 - 11:00";
            resultDataGrid.Rows[6].HeaderCell.Value = "11:00 - 12:00";
            resultDataGrid.Rows[7].HeaderCell.Value = "12:00 - 13:00";
            resultDataGrid.Rows[8].HeaderCell.Value = "13:00 - 14:00";
            resultDataGrid.Rows[9].HeaderCell.Value = "14:00 - 15:00";
            resultDataGrid.Rows[10].HeaderCell.Value = "15:00 - 16:00";
            resultDataGrid.Rows[11].HeaderCell.Value = "16:00 - 17:00";
            resultDataGrid.Rows[12].HeaderCell.Value = "17:00 - 18:00";
            resultDataGrid.Rows[13].HeaderCell.Value = "18:00 - 19:00";
            resultDataGrid.Rows[14].HeaderCell.Value = "19:00 - 20:00";
            resultDataGrid.Rows[15].HeaderCell.Value = "20:00 - 21:00";
            resultDataGrid.Rows[16].HeaderCell.Value = "10:00 - 22:00";
            resultDataGrid.Rows[17].HeaderCell.Value = "Всего";
            resultDataGrid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            totalLabel.Text = String.Format("Всего: {0}", data.Count);
           
        }
        void selectRowsFromDB()
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string sqlExpression = String.Format("SELECT GettingDate FROM Registration WHERE GettingDate >=CONVERT(date, '{0}-{1}-{2} 08:00:00.000') and GettingDate <= CONVERT(date, '{3}-{4}-{5}') {6}", beginDate.Year, beginDate.Month, beginDate.Day, endDate.Year, endDate.Month, endDate.Day, 
                        String.IsNullOrEmpty(procedureName) ? "" : String.Format("and Number='{0}'", procedureName));
                    SqlCommand command = new SqlCommand(sqlExpression, connection);
                    SqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            data.Add(reader.GetDateTime(0));
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Ошибка");
                }
                finally
                {
                    connection.Close();
                }
            }
            for (int i = 0; i < data.Count; i++)
            {
                if (data[i].Hour < 8)
                    data[i] = data[i].AddHours(8 - data[i].Hour);
            }
        }
        void InsertExcel(string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                //// Setting up columns
                DocumentFormat.OpenXml.Spreadsheet.Columns columns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                for (int i = 0; i < resultDataGrid.ColumnCount +1; i++)
                {
                    columns.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Column 
                    {
                        Min = 1,
                        Max = 3,
                        Width = 10,
                        CustomWidth = true
                    });
                   // _sheet.Cells[1, i + 2] = resultDataGrid.Columns[i].HeaderCell.Value.ToString();
                }

                worksheetPart.Worksheet.AppendChild(columns);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = DateTime.Now.ToShortDateString() };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();
                

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                string[] paramsHeaderArray = new string[resultDataGrid.ColumnCount + 1];
                paramsHeaderArray[0] = "";
                for (int j = 0; j < resultDataGrid.ColumnCount; j++)
                {
                    paramsHeaderArray[j+1] = resultDataGrid.Columns[j].HeaderText.ToString();// resultDataGrid.Rows[i].Cells[j].Value.ToString();
                }
                sheetData.AppendChild(ConstructRow(2, paramsHeaderArray));
                //sheetData.AppendChild(ConstructRow(2, paramsHeaderArray));

                for (int i = 0; i < resultDataGrid.RowCount; i++)
                {
                    string[] paramsArray = new string[resultDataGrid.ColumnCount + 1];
                    paramsArray[0] = resultDataGrid.Rows[i].HeaderCell.Value.ToString();
                    for (int j = 0; j < resultDataGrid.ColumnCount ; j++)
                    {

                         paramsArray[j+1] = resultDataGrid.Rows[i].Cells[j].Value.ToString();
                    }
                    sheetData.AppendChild(ConstructRow(2, paramsArray));
                }
                sheetData.AppendChild(ConstructRow(0, "Итого:", data.Count.ToString()));
                worksheetPart.Worksheet.Save();

                
            }
            }
        void InsertStatisticsInExcel(string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                //// Setting up columns
                DocumentFormat.OpenXml.Spreadsheet.Columns columns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                for (int i = 0; i < resultDataGrid.ColumnCount + 1; i++)
                {
                    columns.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Column
                    {
                        Min = 1,
                        Max = 3,
                        Width = 10,
                        CustomWidth = true
                    });
                    // _sheet.Cells[1, i + 2] = resultDataGrid.Columns[i].HeaderCell.Value.ToString();
                }

                worksheetPart.Worksheet.AppendChild(columns);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = DateTime.Now.ToShortDateString() };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();


                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                int startRowIndex = 0, startColumnIndex = 0;
                for (int i = 0; i < daysCount.GetLength(0); i++)
                {
                    bool isNullRow = true;
                    for (int j = 0; j < daysCount.GetLength(1); j++)
                    {
                        if (daysCount[i, j] != 0)
                        {
                            isNullRow = false;
                            break;
                        }
                        
                    }
                    if (!isNullRow)
                    {
                        startRowIndex = i;
                        break;
                    }
                }
                string[] paramsHeaderArray = new string[15];
                paramsHeaderArray[0] = "";
                paramsHeaderArray[1] = "Понедельник";
                paramsHeaderArray[2] = "";
                paramsHeaderArray[3] = "Вторник";
                paramsHeaderArray[4] = "";
                paramsHeaderArray[5] = "Среда";
                paramsHeaderArray[6] = "";
                paramsHeaderArray[7] = "Четверг";
                paramsHeaderArray[8] = "";
                paramsHeaderArray[9] = "Пятница";
                paramsHeaderArray[10] = "";
                paramsHeaderArray[11] = "Суббота";
                paramsHeaderArray[12] = "";
                paramsHeaderArray[13] = "Средняя загрузка службы в течении дня за промежуток времени";
                paramsHeaderArray[14] = "";
                TimeSpan ts = endDate - beginDate;
                double weeks = ts.TotalDays / 7;
                sheetData.AppendChild(ConstructRow(2, paramsHeaderArray));
                int totalCount = 0;
                int[,] totalArray = new int[12 + startRowIndex, daysCount.GetLength(1) + 1];
                for (int i = startRowIndex; i < 12 + startRowIndex; i++)
                {
                    string[] newParamsArray = new string[15];
                    string[] paramsArray = new string[daysCount.GetLength(1) + 3];
                    paramsArray[0] = resultDataGrid.Rows[i-startRowIndex].HeaderCell.Value.ToString();
                    totalCount = 0;
                    int totalWorkerOnWeek = 0;
                    for (int j = 0; j < daysCount.GetLength(1); j++)
                    {
                        try
                        {
                            paramsArray[j + 1] = daysCount[i, j].ToString();
                            totalArray[i, j] = daysCount[i, j];
                            totalCount += daysCount[i, j];
                            totalWorkerOnWeek += returnAreaGrafik()[i - startRowIndex, j ];
                        }
                        catch (Exception eq)
                        {

                        }

                    }
                    
                    newParamsArray[0] = resultDataGrid.Rows[i - startRowIndex].HeaderCell.Value.ToString();
                    newParamsArray[1] = paramsArray[1];
                    newParamsArray[2] = String.Format("{0:f1}", (double.Parse(paramsArray[1]) / (weeks) / returnAreaGrafik()[i - startRowIndex, 0]));//    String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(2), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, 0]);
                    newParamsArray[3] = paramsArray[2];
                    newParamsArray[4] = String.Format("{0:f1}", (double.Parse(paramsArray[2]) / (weeks) / returnAreaGrafik()[i - startRowIndex, 1]));//String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(4), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, 1]);
                    newParamsArray[5] = paramsArray[3];
                    newParamsArray[6] = String.Format("{0:f1}", (double.Parse(paramsArray[3]) /(weeks) / returnAreaGrafik()[i - startRowIndex, 2]));//String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(6), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, 2]);
                    newParamsArray[7] = paramsArray[4];
                    newParamsArray[8] = String.Format("{0:f1}", (double.Parse(paramsArray[4]) / (weeks) / returnAreaGrafik()[i - startRowIndex, 3]));//String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(8), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, 3]);
                    newParamsArray[9] = paramsArray[5];
                    newParamsArray[10] = String.Format("{0:f1}", (double.Parse(paramsArray[5]) /( weeks) / returnAreaGrafik()[i - startRowIndex, 4])); //String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(10), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, 4]);
                    newParamsArray[11] = paramsArray[6];
                    newParamsArray[12] = "";// (double.Parse(paramsArray[6]) / ts.TotalDays / 7 / zav[i - startRowIndex, 5]).ToString();// String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(12), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, 5]);
                    newParamsArray[13] = Math.Round((double)(totalCount / (ts.TotalDays / 7))).ToString();
                    newParamsArray[14] = String.Format("{0:f1}", Math.Round((double)(totalCount / weeks))/totalWorkerOnWeek);

                    //for (int j = 0; j < paramsArray.Length; j++)
                    //{
                    //    if (j % 2 == 0 && j > 1)
                    //    {
                    //        newParamsArray[j + 1] = String.Format("={0}{1}/{2}/{3}", ExcelIntToCharCollums(j), (i - startRowIndex + 2), ts.TotalDays / 7, zav[i - startRowIndex, j]);
                    //    }
                    //    else
                    //    {
                    //        newParamsArray[j + 1] = paramsArray[]
                    //    }

                    //}
                    paramsArray[paramsArray.Length-2] = Math.Round((double)(totalCount/(ts.TotalDays/7))).ToString();
                    paramsArray[paramsArray.Length - 1] = (ts.TotalDays / 7).ToString();
                    sheetData.AppendChild(ConstructRow(2, newParamsArray));
                }
               // sheetData.AppendChild(ConstructRow(0, "Итого:", data.Count.ToString()));
                worksheetPart.Worksheet.Save();


            }
        }
        private void MergeCellsInExcelDoc(int rowNumber, SpreadsheetDocument document, string sheetName)
        {
            DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = GetWorksheet(document, sheetName);
            MergeCells mergeCells;

            if (worksheet.Elements<MergeCells>().Count() > 0)
                mergeCells = worksheet.Elements<MergeCells>().First();
            else
            {
                mergeCells = new MergeCells();

                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                else
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            }

            MergeCell mergeCell = new MergeCell()
            {
                Reference =
                    new StringValue(String.Format("A{0}:C{0}", rowNumber))

            };
            mergeCells.Append(mergeCell);
            worksheet.Save();
        }



        private static DocumentFormat.OpenXml.Spreadsheet.Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook
                .Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart
                .GetPartById(sheets.First().Id);
            return worksheetPart.Worksheet;
        }
        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }
        private Row ConstructRow(uint styleIndex = 0, params string[] values)
        {
            Row resultRow = new Row();
            resultRow.CustomHeight = true;
            double rowHeight = 15.75;
            resultRow.Height = rowHeight;
            bool RowIsIncreased = false;
            foreach (var value in values)
            {
                if (styleIndex != 1 && (int)(value.Length / 100) >= 1)
                {
                    rowHeight += (15.75 * (int)(value.Length / 100));
                    resultRow.Height = rowHeight;
                    RowIsIncreased = true;
                }
                if (styleIndex == 1 && (int)(value.Length / 35) >= 1)
                {
                    rowHeight += (15.75 * (int)(value.Length / 35));
                    resultRow.Height = rowHeight;
                    RowIsIncreased = true;
                }
                if (RowIsIncreased)
                    resultRow.Height = rowHeight;
                resultRow.AppendChild(ConstructCell(value, CellValues.String, styleIndex));
            }
            return resultRow;
        }
        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            DocumentFormat.OpenXml.Spreadsheet.Fonts fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                new DocumentFormat.OpenXml.Spreadsheet.Font( // Index 0 - default
                     new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 12 },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold() { Val = false }),
                new DocumentFormat.OpenXml.Spreadsheet.Font( // Index 1 - header
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 12 },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold() { Val = false },
                    new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = "FFFFFF" }),
                 new DocumentFormat.OpenXml.Spreadsheet.Font( // Index 2 - header
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = 12 },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold() { Val = true },
                    new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = "FFFFFF" })
                );

            Fills fills = new Fills(
                    new DocumentFormat.OpenXml.Spreadsheet.Fill(new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new DocumentFormat.OpenXml.Spreadsheet.Fill(new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new DocumentFormat.OpenXml.Spreadsheet.Fill(new DocumentFormat.OpenXml.Spreadsheet.PatternFill(new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } }) { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            DocumentFormat.OpenXml.Spreadsheet.Borders borders = new DocumentFormat.OpenXml.Spreadsheet.Borders(
                    new DocumentFormat.OpenXml.Spreadsheet.Border(), // index 0 default
                    new DocumentFormat.OpenXml.Spreadsheet.Border( // index 1 black border
                        new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.RightBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.TopBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder()),

                    new DocumentFormat.OpenXml.Spreadsheet.Border( // index 1 black border
                        new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.RightBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.TopBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(new DocumentFormat.OpenXml.Spreadsheet.Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new DocumentFormat.OpenXml.Spreadsheet.CellFormat(new Alignment() { WrapText = true, ShrinkToFit = true }), // default
                    new DocumentFormat.OpenXml.Spreadsheet.CellFormat(new Alignment() { WrapText = true, ShrinkToFit = true }) { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true },
                    new DocumentFormat.OpenXml.Spreadsheet.CellFormat(new Alignment() { WrapText = true, ShrinkToFit = true }) { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true } // split cells
                );
            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);
            return styleSheet;
        }


        private void InToExcelButton_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            
            
            //Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            //try
            //{
            //ExcelApp.Application.Workbooks.Add(Type.Missing);
            //ExcelApp.Columns.ColumnWidth = 15;
            //Microsoft.Office.Interop.Excel.Workbook workbook = ExcelApp.Workbooks[1];
            //Microsoft.Office.Interop.Excel.Worksheet _sheet;  _sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);


            //var dgw1 = new DataGridView();
            //dgw1 = resultDataGrid;
            //var cntColl = resultDataGrid.ColumnCount;
            //var cntrow = resultDataGrid.RowCount;

          
            //    if (dgw1.RowCount != 0)
            //    {

            //        //Заполнение заголовков столбцов
            //        //for (int coll = 1; coll <= cntColl; coll++)
            //        //{
            //        //    _sheet.Cells[1, coll] = dgw1.Columns[coll - 1].HeaderCell.Value;
            //        //    _sheet.Range[ExcelIntToCharCollums(coll).ToString() + "1"].ColumnWidth = 15; //Ширина столбца
            //        //    _sheet.Range[ExcelIntToCharCollums(coll).ToString() + "1"].WrapText = true; //Перенос текста 
            //        //    //делаем их жирными
            //        //    _sheet.Range[ExcelIntToCharCollums(coll).ToString() + "1"].Font.Bold = true; //Жирный шрифт
            //        //    _sheet.Range[ExcelIntToCharCollums(coll).ToString() + "1"].Font.Size = 14; // Размер шрифта
            //        //}
            //        //Заполнение ячеек данными
            //        for (int i = 0; i < resultDataGrid.ColumnCount; i++)
            //        {
            //            _sheet.Cells[1, i+2] = resultDataGrid.Columns[i].HeaderCell.Value.ToString();
            //        }
            //        for (int i = 0; i < resultDataGrid.RowCount; i++)
            //        {
            //            _sheet.Cells[i + 2, 1] = resultDataGrid.Rows[i].HeaderCell.Value.ToString();
            //        }
            //        for (int row = 0; row <= cntrow - 1; row++)
            //        {
            //            for (int coll = 2; coll <= cntColl+1; coll++)
            //            {
            //                _sheet.Cells[row + 2, coll] = dgw1.Rows[row].Cells[coll - 2].Value;
            //               // _sheet.Range[ExcelIntToCharCollums(coll).ToString() + row + 1.ToString()].WrapText = true; //Перенос текста 
            //            }
            //        }
            //        _sheet.Cells[cntrow + 2, 1] = selectedPeriod.Text;
            //        _sheet.Cells[cntrow + 3, 1] = totalLabel.Text;
                  
            //        Microsoft.Office.Interop.Excel.Range _excelCells = (Microsoft.Office.Interop.Excel.Range)_sheet.get_Range("A" + (cntrow + 2), "C" + (cntrow + 2)).Cells;

            //        // Производим объединение
            //        _excelCells.Merge(Type.Missing);
            //        ExcelApp.Visible = true;
            //        // Уничтожение объекта Excel.
            //        Marshal.ReleaseComObject(ExcelApp);
            //        // Вызываем сборщик мусора для немедленной очистки памяти
            //        GC.GetTotalMemory(true);
            //    }

            //}
            //catch (Exception ee)
            //{ MessageBox.Show(ee.ToString()); }
            //finally
            //{
            //    Marshal.ReleaseComObject(ExcelApp);
            //}
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            InsertExcel(saveFileDialog1.FileName);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            saveFileDialog2.ShowDialog();
            ////  getStatistics();
            //string resRow = "";
            //for (int i = 0; i < daysCount.GetLength(0); i++)
            //{
            //    for (int j = 0; j < daysCount.GetLength(1); j++)
            //    {
            //        resRow = resRow + daysCount[i, j].ToString() + " ";
            //    }
            //    resRow = resRow + "\n";
            //}
            //MessageBox.Show(resRow);
        }

        private void saveFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            InsertStatisticsInExcel(saveFileDialog2.FileName);
        }

        private string ExcelIntToCharCollums(int num)
        {
            switch (num)
            {
                case 1: return "A";
                case 2: return "B";
                case 3: return "C";
                case 4: return "D";
                case 5: return "E";
                case 6: return "F";
                case 7: return "G";
                case 8: return "H";
                case 9: return "I";
                case 10: return "J";
                case 11: return "K";
                case 12: return "L";
                case 13: return "M";
                case 14: return "N";
                case 15: return "O";
                case 16: return "P";
                case 17: return "Q";
                case 18: return "R";
                case 19: return "S";
                case 20: return "T";
                case 21: return "U";
                case 22: return "V";
                case 23: return "W";


            }
            return null;
        }
        int[,] returnAreaGrafik()
        {
            switch (thisArea)
            {

                case "zav": return zav; 
                case "len": return len; 
                case "mingor": return mingor; 
                case "mos": return mos; 
                case "okt": return okt; 
                case "par": return per;
                case "per": return par; 
                case "sov": return sov;
                case "cen": return cen; 
                case "frun": return frun;
                default: return zav;
            }
        }
      
        int[,] per = //Готово
{
            {1,1,1,1,1,0 },//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 6,6,6,6,6,2},//11
            { 6,6,6,6,6,2},//12
            { 2,2,2,2,2,2},//13
            { 4,4,4,4,4,0},//14
            { 6,6,6,6,6,0},//15
            { 5,5,5,5,5,0},//16
            { 5,5,5,5,5,0},//17
            { 2,2,2,2,2,0},//18
            { 2,2,2,2,2,0},//19
        };
        int[,] par = //Готово
{
            { 2,2,2,2,2,0},//8
            { 3,3,3,3,3,1},//9
            { 3,3,3,3,3,1},//10
            { 5,5,5,5,5,1},//11
            { 5,5,5,5,5,1},//12
            { 2,2,2,2,2,1},//13
            { 3,3,3,3,3,0},//14
            { 5,5,5,5,5,0},//15
            { 5,5,5,5,5,0},//16
            { 5,5,5,5,5,0},//17
            { 2,2,2,2,2,0},//18
            { 2,2,2,2,2,0},//19
        };
        int[,] len = //Готово
{
            { 3,3,3,3,3,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 7,7,7,7,7,2},//11
            { 4,5,5,5,4,2},//12
            { 4,5,5,5,4,2},//13
            { 5,6,6,6,5,0},//14
            { 4,5,5,5,4,0},//15
            { 6,7,7,7,6,0},//16
            { 4,4,4,4,4,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] zav = //Готово
        {
            { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] mos =
{
              { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] cen =
       {
            { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] mingor =
       {
            { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] sov =
       {
            { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] okt =
       {
            { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
        int[,] frun =
       {
            { 1,1,1,1,1,0},//8
            { 4,4,4,4,4,2},//9
            { 4,4,4,4,4,2},//10
            { 8,8,8,8,8,2},//11
            { 7,7,7,7,7,2},//12
            { 3,3,3,3,3,2},//13
            { 5,5,5,5,5,0},//14
            { 8,8,8,8,8,0},//15
            { 8,8,8,8,8,0},//16
            { 7,7,7,7,7,0},//17
            { 3,3,3,3,3,0},//18
            { 3,3,3,3,3,0},//19
        };
    }
}
