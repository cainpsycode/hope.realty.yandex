using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;
using NLog;

namespace xlsx.convert.yrl
{
    public class ConvertExcelToYrl
    {
        private Logger _logger;
        private SalesAgent _salesAgent;

        public ConvertExcelToYrl(Logger logger, SalesAgent salesAgent)
        {
            _logger = logger;
            _salesAgent = salesAgent;
        }

        private static string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        /// <summary>
        ///  Read Data from selected excel file into DataTable
        /// </summary>
        /// <param name="filename">Excel File Path</param>
        /// <param name="skipRows"></param>
        /// <returns></returns>
        private DataTable ReadExcelFile(string filename, int skipRows)
        {
            // Initialize an instance of DataTable
            DataTable dt = new DataTable();

            try
            {
                // Use SpreadSheetDocument class of Open XML SDK to open excel file
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
                {
                    // Get Workbook Part of Spread Sheet Document
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                    // Get all sheets in spread sheet document 
                    IEnumerable<Sheet> sheetcollection = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                    // Get relationship Id
                    string relationshipId = sheetcollection.First().Id.Value;

                    // Get sheet1 Part of Spread Sheet Document
                    WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);

                    // Get Data in Excel file
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    IEnumerable<Row> rowcollection = sheetData.Descendants<Row>();

                    if (rowcollection.Count() == 0)
                    {
                        return dt;
                    }

                    // Add columns
                    int index = 0;
                    var strIndex = Enumerable.Range(1, 100).Select(ExcelColumnFromNumber).ToList();
                    foreach (Cell cell in rowcollection.ElementAt(skipRows))
                    {
                        while (!cell.CellReference.Value.StartsWith(strIndex.ElementAt(index), StringComparison.InvariantCultureIgnoreCase))
                        {
                            dt.Columns.Add("");
                            index++;
                        }

                        dt.Columns.Add(GetValueOfCell(spreadsheetDocument.WorkbookPart, cell));
                        index++;
                    }

                    // Add rows into DataTable
                    foreach (Row row in rowcollection.Skip(skipRows))
                    {
                        DataRow temprow = dt.NewRow();
                        int columnIndex = 0;

                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            // Get Cell Column Index
                            int cellColumnIndex = GetColumnIndex(GetColumnName(cell.CellReference));

                            if (columnIndex < cellColumnIndex)
                            {
                                do
                                {
                                    temprow[columnIndex] = string.Empty;
                                    columnIndex++;
                                }

                                while (columnIndex < cellColumnIndex);
                            }
                            
                            if (columnIndex < dt.Columns.Count)
                            {
                                temprow[columnIndex] = GetValueOfCell(spreadsheetDocument.WorkbookPart, cell);
                            }
                                
                            columnIndex++;
                        }

                        // Add the row to DataTable
                        // the rows include header row
                        dt.Rows.Add(temprow);
                    }
                }

                // Here remove header row
                dt.Rows.RemoveAt(0);
                return dt;
            }
            catch (IOException ex)
            {
                throw new IOException(ex.Message);
            }
        }

        private enum Formats
        {
            General = 0,
            Number = 1,
            Decimal = 2,
            Currency = 164,
            Accounting = 44,
            DateShort = 14,
            DateLong = 165,
            Time = 166,
            Percentage = 10,
            Fraction = 12,
            Scientific = 11,
            Text = 49
        }

        private static string GetValueOfCell(WorkbookPart workbookPart, Cell cell)
        {
            if (cell == null)
            {
                return null;
            }

            string value = "";
            if (cell.DataType == null) // number & dates
            {
                
                int styleIndex = (int) (cell.StyleIndex.HasValue ? cell.StyleIndex.Value : 0);
                CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);
                uint formatId = cellFormat.NumberFormatId.Value;

                if (formatId == (uint)Formats.DateShort || formatId == (uint)Formats.DateLong)
                {
                    double oaDate;
                    if (double.TryParse(cell.InnerText, out oaDate))
                    {
                        value = DateTime.FromOADate(oaDate).ToShortDateString();
                    }
                }
                else
                {
                    value = cell.InnerText;
                }
            }
            else // Shared string or boolean
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.InnerText));
                        value = ssi.Text.Text;
                        break;
                    case CellValues.Boolean:
                        value = cell.CellValue.InnerText == "0" ? "false" : "true";
                        break;
                    default:
                        value = cell.CellValue.InnerText;
                        break;
                }
            }

            return value;
        }

        private static DateTime? ParseExcelDateTime(string value)
        {
            double oaDateAsDouble;
            if (!double.TryParse(value, out oaDateAsDouble)) return null;
            return DateTime.FromOADate(oaDateAsDouble);
        }

        /// <summary>
        /// Get Column Name From given cell name
        /// </summary>
        /// <param name="cellReference">Cell Name(For example,A1)</param>
        /// <returns>Column Name(For example, A)</returns>
        private string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name of cell
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        /// <summary>
        /// Get Index of Column from given column name
        /// </summary>
        /// <param name="columnName">Column Name(For Example,A or AA)</param>
        /// <returns>Column Index</returns>
        private int GetColumnIndex(string columnName)
        {
            int columnIndex = 0;
            int factor = 1;

            // From right to left
            for (int position = columnName.Length - 1; position >= 0; position--)   
            {
                // For letters
                if (Char.IsLetter(columnName[position]))
                {
                    columnIndex += factor * ((columnName[position] - 'A') + 1) - 1;
                    factor *= 26;
                }
            }

            return columnIndex;
        }

        private XDocument ClearNs(XDocument doc)
        {
            foreach (var node in doc.Root.Descendants().Where(n => n.Name.NamespaceName == ""))
            {
                // Remove the xmlns='' attribute. Note the use of
                // Attributes rather than Attribute, in case the
                // attribute doesn't exist (which it might not if we'd
                // created the document "manually" instead of loading
                // it from a file.)
                node.Attributes("xmlns").Remove();
                // Inherit the parent namespace instead
                node.Name = node.Parent.Name.Namespace + node.Name.LocalName;
            }

            return doc;
        }

        /// <summary>
        /// Convert DataTable to Xml format
        /// </summary>
        /// <param name="filename">Excel File Path</param>
        /// <param name="skipFirstRow">Skip first row with date</param>
        /// <returns>Xml format string</returns>
        public XDocument GetXml(string filename, int skipRows)
        {
            DataTable table = ReadExcelFile(filename, skipRows);
            var offers = table.Rows.Cast<DataRow>().Select((i, n) =>
            {
                try
                {
                    var offer = ParseOfferSaleNew(i);
                    return offer;
                }
                catch (Exception ex)
                {
                    _logger.Error("Ошибка обработки строки {0}: {1}", n, ex);
                    return null;
                }
            }).Where(i => i != null);

            XNamespace ns = "http://webmaster.yandex.ru/schemas/feed/realty/2010-06";

            var xDocument = new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XElement(ns + "realty-feed",
                    new XElement("generation-date", DateTime.Now.ToIso8601()),
                    offers));

            return ClearNs(xDocument);
        }

        private string[] PermDistricts => new[] {
            "дзержинский", "индустриальный", "кировский", "ленинский", "мотовилихинский", "новые ляды",
            "орджоникидзевский", "свердловский"
        };

        private static Dictionary<string, string> XmlEscapedChars => new Dictionary<string, string>()
        {
            {"\"", "&quot;" },
            {"&", "&amp;" },
            {">", "&gt;" },
            {"<", "&lt;" },
            {"'", "&apos;" }
        };

        private static Dictionary<string, string> BuildingState => new Dictionary<string, string>()
        {
            {"дом построен", "built" },
            {"сдан в эксплуатацию", "hand-over" },
            {"строится", "unfinished" }
        };

        private string EncodeYrlText(string text)
        {
            return text == null ? null : XmlEscapedChars.Aggregate(text, (current, esc) => current.Replace(esc.Key, esc.Value));
        }

        private XElement ParseOfferSaleNew(DataRow row)
        {
            Func<ColumnIndex, string> rowValue = col =>
            {
                try { return EncodeYrlText(row.ItemArray[(int) col] as string);  }
                catch (Exception ex) { throw new Exception($"Ошибка обработки ячейки {col}", ex); }
            };
            Func<string, string, XElement> buildSpace = (name, value) => new XElement(name, new XElement("value", value), new XElement("unit", "кв. м"));
            Func<string, ColumnIndex, XElement> buildSpaceColumn = (name, col) => string.IsNullOrEmpty(rowValue(col)) ? null : buildSpace(name, rowValue(col));

            Match matchAgent = new Regex(@"^(.*)\s+(.*)$").Match((string)row.ItemArray[(int)ColumnIndex.Agent]);
            string phone = new Regex(@"^\+?8").Replace(matchAgent.Groups[1].Value, "+7");
            string agent = matchAgent.Groups[2].Value;

            float[] roomsArea = rowValue(ColumnIndex.RoomsArea).Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries).Select(float.Parse).ToArray();
            if (roomsArea.Length > 0 && (roomsArea.Length != int.Parse(rowValue(ColumnIndex.Rooms))))
            {
                throw new Exception("Количество перечисленных площадей комнат не соответствует количеству комнат");
            };

            var offer = new XElement(
                "offer",
                new XAttribute("internal-id", int.Parse(rowValue(ColumnIndex.Id))),
                new XElement("type", "продажа"),
                new XElement("property-type", "жилая"),
                new XElement("category", rowValue(ColumnIndex.Category)),
                new XElement("url", rowValue(ColumnIndex.Url)),
                new XElement("creation-date", DateTime.Parse(rowValue(ColumnIndex.CreationDate)).ToIso8601()),
                new XElement("last-update-date", DateTime.Parse(rowValue(ColumnIndex.LastUpdateDate)).ToIso8601()),
                new XElement("location", new XElement[] {
                    new XElement("country", "Россия"),
                    new XElement("region", "Пермский край"),
                    new XElement("district", rowValue(ColumnIndex.District)),
                    new XElement("locality-name", rowValue(ColumnIndex.LocalityName)),
                    new XElement("sub-locality-name", rowValue(ColumnIndex.SubLocalityName)),
                    new XElement("address", rowValue(ColumnIndex.Address))
                }.Where(i => !string.IsNullOrEmpty(i.Value))
                ),
                new XElement("sales-agent",
                    new XElement("phone", phone),
                    new XElement("name", agent),
                    new XElement("category", "агентство"),
                    new XElement("organization", _salesAgent.Organization),
                    new XElement("url", _salesAgent.Url),
                    new XElement("photo", _salesAgent.Photo)
                ),
                new XElement("deal-status", rowValue(ColumnIndex.DealStatus)),
                new XElement("price",
                    new XElement("value", decimal.Parse(rowValue(ColumnIndex.Price)) * 1000m),
                    new XElement("currency", "RUR")
                ),
//                new XElement("image", null),
                buildSpace("area", rowValue(ColumnIndex.Area)),
                new XElement[]
                {
                    buildSpaceColumn("living-space", ColumnIndex.Living),
                    buildSpaceColumn("kitchen-space", ColumnIndex.Kitchen)
                }
                .Concat(roomsArea.Select(i => buildSpace("room-space", i.ToString(CultureInfo.InvariantCulture))))
                .Where(i => i != null),
                new XElement("renovation", rowValue(ColumnIndex.Renovation)),
                new XElement("description", rowValue(ColumnIndex.Description)),
                new XElement("new-flat", "да"),
                new XElement("floor", rowValue(ColumnIndex.Floor)),
                new XElement("rooms", rowValue(ColumnIndex.Rooms)),
                new XElement[] {
                    new XElement("studio", rowValue(ColumnIndex.Studio)),
                    new XElement("open-plan", rowValue(ColumnIndex.OpenPlan)),
                    new XElement("balcony", rowValue(ColumnIndex.Balcony))
                }.Where(i => !string.IsNullOrEmpty(i.Value)),
                new XElement("floors-total", rowValue(ColumnIndex.FloorsTotal)),
//                new XElement("building-name", rowValue(ColumnIndex.BuildingName)),
//                new XElement("yandex-building-id", rowValue(ColumnIndex.YandexBuildingId)),
                new XElement("building-state", BuildingState[rowValue(ColumnIndex.BuildingState)]),
                new XElement("built-year", rowValue(ColumnIndex.BuildingYear)),
                new XElement("ready-quarter", rowValue(ColumnIndex.ReadyQuarter)),
                new XElement("building-type", rowValue(ColumnIndex.BuildingType))
            );

            return offer;
        }


//        private OfferSaleNew ParseOfferSaleNew(DataRow row)
//        {
//            OfferLocation location = new OfferLocation()
//            {
//                Country = "Россия",
//                Region = "Пермский край",
//                District = (string)row.ItemArray[(int)ColumnIndex.District],
//                LocalityName = (string)row.ItemArray[(int)ColumnIndex.LocalityName],
//                SubLocalityName = (string)row.ItemArray[(int)ColumnIndex.SubLocalityName],
//                Address = (string)row.ItemArray[(int)ColumnIndex.Address]
//            };
//
//            Regex regex = new Regex("^(.*)\\s+(.*)$");
//            Match match = regex.Match((string)row.ItemArray[(int)ColumnIndex.Agent]);
//            Regex regEx = new Regex("^\\+?8");
//            string phone = regEx.Replace(match.Groups[1].Value, "+7");
//
//            var salesAgent = new SalesAgent(
//                match.Groups[2].Value,
//                _salesAgent.Organization,
//                _salesAgent.Phones.Concat(new[] { phone }).ToArray(),
//                _salesAgent.Category,
//                _salesAgent.Url,
//                _salesAgent.Email,
//                _salesAgent.Photo);
//
//            var offer = new OfferSaleNew(
//                salesAgent,
//                int.Parse((string)row.ItemArray[(int)ColumnIndex.Id]),
//                OfferSaleNew.OfferType.Sale,
//                OfferSaleNew.PropertyType.Living,
//                EnumByDescription((string)row.ItemArray[(int)ColumnIndex.Category]),
//                DateTime.Parse((string)row.ItemArray[(int)ColumnIndex.CreationDate]),
//                DateTime.Parse((string)row.ItemArray[(int)ColumnIndex.LastUpdateDate]),
//                location,
//                (string)row.ItemArray[(int)ColumnIndex.Url],
//                OfferSaleNew.DealStatusType.SaleDeveloper,
//                Decimal.Parse((string)row.ItemArray[(int)ColumnIndex.Price]) * 1000);
//
//            return offer;
//        }
//
//        private XElement GetElementOffer(OfferSaleNew offer)
//        {
//            return new XElement(
//                "offer",
//                new XAttribute("internal-id", offer.Id),
//                new XElement("type", offer.Type.GetDescription()),
//                new XElement("property-type", offer.Property.GetDescription()),
//                new XElement("category", offer.Category.GetDescription()),
//                new XElement("url", offer.Url),
//                new XElement("creation-date", offer.CreationDate.ToIso8601()),
//                new XElement("last-update-date", offer.LastUpdateDate.ToIso8601()),
//                new XElement("location", new XElement[] {
//                    new XElement("country", offer.LocationPoint.Country),
//                    new XElement("region", offer.LocationPoint.Region),
//                    new XElement("district", offer.LocationPoint.District),
//                    new XElement("locality-name", offer.LocationPoint.LocalityName),
//                    new XElement("sub-locality-name", offer.LocationPoint.SubLocalityName),
//                    new XElement("address", offer.LocationPoint.Address)
//                }.Where(i => !string.IsNullOrEmpty(i.Value))
//                ),
//                new XElement("sales-agent",
//                    offer.Agent.Phones.Select(i => new XElement("phone", i)),
//                    new XElement("category", offer.Agent.Category.GetDescription()),
//                    new XElement("organization", offer.Agent.Organization),
//                    new XElement("url", offer.Agent.Url),
//                    new XElement("photo", offer.Agent.Url)
//                ),
//                new XElement("deal-status", offer.DealStatus.GetDescription()),
//                new XElement("price",
//                    new XElement("value", offer.Price),
//                    new XElement("currency", "RUR")
//                ),
//                new XElement("image", null)
//            );
//        }

        private static OfferSaleNew.CategoryType EnumByDescription(string description)
        {
            return (OfferSaleNew.CategoryType)
                    typeof (OfferSaleNew.CategoryType).GetEnumValueByAttribute(typeof (DescriptionAttribute),
                        i => ((DescriptionAttribute) i).Text.Equals(description.ToLower()));
        }

        private enum ColumnIndex : int
        {
            Id = 0,
            Category,
            Url,
            CreationDate,
            LastUpdateDate,
            District,
            LocalityName,
            SubLocalityName,
            Address,
            DealStatus,
            Price,
            Floor,
            FloorsTotal,
            Rooms,
            Area,
            Living,
            Kitchen,
            RoomsArea,
            Renovation,
            Studio,
            OpenPlan,
            Balcony,
            BuildingName,
            BuildingState,
            BuildingYear,
            ReadyQuarter,
            BuildingType,
            Agent,
            Description
        }
    }


    
}
