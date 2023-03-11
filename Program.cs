using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Text.RegularExpressions;

class Program
{
    static void Main(string[] args)
    {
        string[] arg = { AppDomain.CurrentDomain.BaseDirectory + "MODEL.xlsx", "000000", "0" };
        if(args.Length>0){
            for (int index = 0; index < args.Length; index++)
                arg[index] = args[index];
        }
        Console.WriteLine("{0} : {1} : {2}", arg[0], arg[1], arg[2]);
        ReadXLSX(arg[0], arg[1], arg[2]);
    }

    private static void ReadXLSX(string filePath, string codigoSKU, string cantidad)
    {
        var ini = DateTime.Now;
        try
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, true))
            {
                if (document != null)
                {
                    WorkbookPart? workbookPart = document.WorkbookPart;
                    if (workbookPart != null)
                    {

                        string IndexWorkSheet = GetIDSheet(workbookPart, "Inventario");
                        if (!string.IsNullOrEmpty(IndexWorkSheet))
                        {
                            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(IndexWorkSheet);

                            Worksheet worksheet = worksheetPart.Worksheet;
                            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();

                            if (sheetData != null)
                            {
                                var (IndexRowPedido, NameCellSelectPedido) = SeekCellIndex(sheetData, workbookPart, "PEDIDO");

                                if (IndexRowPedido > 0)
                                {
                                    var (IndexRowRestos, NameCellSelectRestos) = SeekCellIndex(sheetData, workbookPart, "CODIGO");
                                    if (IndexRowRestos > 0)
                                    {
                                        var (IndexRowWriteCell, NameCellWrite) = ReadValueCell(sheetData, workbookPart, NameCellSelectRestos, codigoSKU);
                                        if (IndexRowWriteCell > 0)
                                            WriteValueCell(sheetData, IndexRowWriteCell, NameCellSelectPedido, cantidad);
                                    }

                                    worksheetPart.Worksheet.Save();
                                    workbookPart.Workbook.Save();
                                    document.Save();
                                }

                            }
                        }


                    }

                }
            }
        }
        catch (IOException)
        {
            // El archivo está actualmente abierto, no puede continuar.
            Console.WriteLine("No se puede abrir");
        }

        Console.WriteLine("{0} : {1}", ini, DateTime.Now);
        // return Task.CompletedTask;
    }

    private static string GetIDSheet(WorkbookPart workbookPart, string keywordsearh)
    {
        string sheetName = string.Empty;
        Workbook workbook = workbookPart.Workbook;
        Sheets? sheets = workbookPart.Workbook.Sheets;
        if (sheets != null)
            foreach (Sheet sheet in sheets)
            {
                sheetName = string.Format("{0}", sheet.Name);
                if (sheetName.ToString().Trim().ToUpper().Equals(keywordsearh.Trim().ToUpper()))
                {
                    Console.WriteLine("El nombre de la hoja es: " + sheet.Name);
                    return string.Format("{0}", sheet.Id);
                }

            }
        return string.Empty;
    }
    private static (uint, string) SeekCellIndex(SheetData sheetData, WorkbookPart workbookPart, string KeyWordSeek)
    {
        uint IndexRow = 0;
        string NameCellSelect = string.Empty;
        foreach (Row row in sheetData.Elements<Row>())
        {
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (GetCellValue(cell, workbookPart).Trim().ToUpper().Equals(KeyWordSeek.Trim().ToUpper()))
                {
                    IndexRow = row.RowIndex?.Value ?? 0;
                    NameCellSelect = GetColumnNameFromCellReference(cell.CellReference);
                    // Console.WriteLine("{0}  : {1} : {2} : {3} : {4}", cell.CellReference.Value, NameCellSelect + IndexRow, IndexRow, row.RowIndex, GetCellValue(cell, workbookPart));
                    break;
                }

            }
            if (IndexRow > 0)
            {
                break;
            }
        }

        return (IndexRow, NameCellSelect);
    }
    private static (uint, string) ReadValueCell(SheetData sheetData, WorkbookPart workbookPart, string NameCellSelect, string keyWordSeek)
    {
        uint IndexRowSeek = 0;
        string NameCellSeek = string.Empty;
        foreach (Row row in sheetData.Elements<Row>())
        {
            Cell? cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == NameCellSelect + row.RowIndex);
            if (cell != null)
            {

                if (GetCellValue(cell, workbookPart).Trim().ToUpper().Equals(keyWordSeek.Trim().ToUpper()))
                {
                    IndexRowSeek = row.RowIndex?.Value ?? 0;
                    NameCellSeek = GetColumnNameFromCellReference(cell.CellReference);
                    return (IndexRowSeek, NameCellSeek);
                }
            }
        }
        return (IndexRowSeek, NameCellSeek);
    }

    private static void WriteValueCell(SheetData sheetData, uint IndexRow, string NameCellSelect, string NewValueCell)
    {
        Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == IndexRow);
        if (row != null)
        {
            // int newValueCell = new Random().Next(1, 201);
            // Console.WriteLine("row: {0}", row.RowIndex);
            Cell? cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == NameCellSelect + IndexRow);
            if (cell != null)
            {
                CellValue cellValue = new CellValue(NewValueCell.ToString());
                cell.CellValue = cellValue;
            }
            else
            {
                cell = new Cell() { CellReference = NameCellSelect + IndexRow, DataType = CellValues.String };
                row.InsertAt(cell, (int)IndexRow);
                CellValue cellValue = new CellValue(NewValueCell.ToString());
                cell.CellValue = cellValue;
            }
        }

    }

    private static string GetColumnNameFromCellReference(string? cellReference)
    {
        // Obtener solo la parte de la referencia de celda que contiene las letras (columna)
        if (cellReference == null)
        {
            return string.Empty;
        }

        return Regex.Replace(cellReference, @"[^a-zA-Z]", "");
    }

    static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell == null || cell.CellValue == null)
        {
            return String.Empty;
        }

        if (cell.DataType == null)
        {
            return cell.CellValue.Text;
        }

        switch (cell.DataType.Value)
        {
            case CellValues.SharedString:
                if (workbookPart?.SharedStringTablePart?.SharedStringTable != null)
                {
                    var stringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                    string cellValueText = cell.CellValue.Text;
                    int index = int.Parse(cellValueText);
                    if (stringTable != null && index < stringTable.Count())
                    {
                        OpenXmlElement? element = stringTable.ElementAtOrDefault(index);
                        if (element != null)
                        {
                            return element.InnerText;
                        }
                    }
                }
                break;
            default:
                return cell.CellValue.Text;
        }

        return String.Empty;
    }

    private void WriteXLSX(string filePath)
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            Worksheet worksheet = worksheetPart.Worksheet;
            if (worksheet != null)
            {
                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData != null)
                {
                    // Escribir datos en las celdas
                    for (int i = 1; i <= 5; i++)
                    {
                        Row row = new Row();
                        for (int j = 1; j <= 3; j++)
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(string.Format("Valor {0},{1}", i, j));
                            row.Append(cell);
                        }
                        sheetData.Append(row);
                    }
                }
            }

            workbookPart.Workbook.Save();
        }
    }
}



