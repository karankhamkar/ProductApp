using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace ProductApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("====================== Welcome to the Product Store =======================\n");
            List<Product> productList = new List<Product>();
            int option;
            char again;
            do
            {
                Menu();
                option = int.Parse(Console.ReadLine());
                switch (option) 
                {
                    case 1:
                        AddProduct();
                        break;

                    case 2:
                        DeleProduct();
                        break;

                    case 3:
                        UpdateProduct();
                        break;

                    case 4:
                        TabularFormat();
                        break;

                    case 5:
                        ConvertToExcelFile();
                        break;

                    case 6:
                        Console.WriteLine("Show import Excel sheet.");
                        productList = ImportFromExcel();
                        TabularFormat();
                        Console.WriteLine($"file has been succesfully Imported ");
                        break;
                }
                Console.WriteLine("Do you want to continue, press y / n : ");
                
            }
            while (option < 7);
            Console.WriteLine();

            #region Non_Static Methods

            void AddProduct()
            {
                do
                {
                    string name = GetName();
                    int brandId = GetBrandId();
                    DateTime mfgDate = GetManufactureDate();

                    Product product = new Product(name, brandId, mfgDate);
                    productList.Add(product);

                    Console.WriteLine("Do you want to continue, press y / n : ");
                    again = char.Parse(Console.ReadLine());
                    Console.WriteLine();
                }
                while (again == 'y' || again == 'Y');
                foreach (var item in productList)
                {
                    DisplayProducts(item);
                }

            }

            void DeleProduct()
            {
                Console.WriteLine("Delete Product");
                Console.Write("Enter Product Id Which You Want To Delete : ");
                int productid = int.Parse(Console.ReadLine());
                Delete(productid);
                Console.WriteLine("Product Deleted Succesfully");
            }

            void UpdateProduct()
            {
                Console.WriteLine("Enter Product Id Which You Want To Update.");
                int productId = int.Parse(Console.ReadLine());
                Update(productId);
                Console.Clear();
                Console.WriteLine("Product Updated Succesfully");
            }

            void TabularFormat()
            {
                var rowFormat = "| {0,-10} | {1,-20} | {2,-10} | {3,-25} | {4,-17} | {5,-17} |";

                Console.WriteLine("+---------------------------------------------------------------------------------------------------------------------+");
                Console.WriteLine(rowFormat, "Product Id", "Name", "Brand Id", "ManufacuringDate", "ManufacturingMonth", "ManufacturingYear");
                Console.WriteLine("+---------------------------------------------------------------------------------------------------------------------+");

                foreach (Product product in productList)
                {
                    Console.WriteLine(rowFormat, product.ProductId, product.Name, product.BrandId.ToString(), product.MfgDate.ToString(), product.ManufactureMonth, product.ManufactureYear.ToString());
                }

                Console.WriteLine("+---------------------------------------------------------------------------------------------------------------------+");
            }

            void ConvertToExcelFile()
            {
                Console.WriteLine("Where do you want to store?\n1 for Dextop 2 for Document");
                int userchoice = int.Parse(Console.ReadLine());
                string path = ExportToExcel(userchoice, productList);
                Console.WriteLine($"file has been succesfully exported at {path}");
            }

            void Delete(int productid)
            {
                Product product = Search(productid);

                if (product != null)
                {
                    productList.Remove(product);
                }
            }
            Product Search(int productid)
            {
                foreach (Product item in productList)
                {
                    if (item.ProductId == productid)
                    {
                        return item;
                    }
                }
                return null;
            }
            void Update(int productId)
            {
                Product product = Search(productId);
                if (product != null)
                {
                    Console.WriteLine("Enter Updated Values For The Product : ");
                    string name = GetName();
                    int brandid = GetBrandId();
                    DateTime manufacturingdate = GetManufactureDate();

                    product.Name = name;
                    product.BrandId = brandid;
                    product.MfgDate = manufacturingdate;

                    Console.WriteLine("Product Updated Successfully.");
                }
            }
            #endregion

        }




        #region Static Methods
        private static void Menu()
        {
            Console.WriteLine("Select your choice from following Menu : \n1. Add.\n2. Delete.\n3. Update\n4. Show List in Tabular Format.\n5. Export To Excel.\n6. Import From Excel.\n7. Exit Program.\nEnter Your Choice : ");
        }

        private static string GetName()
        {
            Console.WriteLine("Enter Name : ");
            string input = Console.ReadLine();
            if(Product.IsNameValid(input))
            {
                return input;
            }
            else
            {
                Console.WriteLine("Product name exceeds the character limit.\nPlease Re-enter name.");
                return GetName();
            }
        }
        private static int GetBrandId()
        {
            Console.WriteLine("Enter Brand Id : ");
            string input = Console.ReadLine();
            int brandId;
            bool isValidBrand = Brand.IsValidBrand(input, out brandId);
            if(isValidBrand)
            {
                return brandId;
            }
            else
            {
                Console.WriteLine("Invalid Brand Id.\nRe-enter Band Id.");
               return GetBrandId();
            }
        }
        private static DateTime GetManufactureDate()
        {
            Console.WriteLine("Enter Manufacture Date(MM-DD-YYYY) : ");
            string input = Console.ReadLine();
            DateTime mfgDate;
            bool isValidDate = DateTime.TryParse(input, out mfgDate);
            if (isValidDate)
            {
                if (mfgDate < DateTime.Now)
                {
                    return mfgDate;
                }
                else
                {
                    Console.WriteLine("Future date not Acceptable.\nPlease Re-enter Date.");
                    GetManufactureDate();
                }
            }
            else
            {
                Console.WriteLine("Invalid Date.\nPlease Re-enter Date.");
            }
            return GetManufactureDate();
        }
        public static void DisplayProducts(Product product)
        {
            if (product.Name.Length > 10)
            {
                product.Name = product.Name.Substring(0, 7) + "...";
            }
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"\nProduct ID : {product.ProductId}\nProduct Name : {product.Name}\nProduct Brand ID : {product.BrandId}\nProduct Manufacture Date : {product.MfgDate}\nProduct manufacture Year : {product.ManufactureYear}\nProduct manufacture Month : {product.ManufactureMonth}\nProduct was Expire at : {product.IsExpire}\nProduct Best Before : {product.IsExpiringDate}");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("\n♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥♥");
            Console.ForegroundColor = ConsoleColor.White;
        }

        #region ExportToExcel Methods

        private static string ExportToExcel(int userchoice, List<Product> productList)
        {
            string OutPutFileDirectory = string.Empty;
            if (userchoice == 1)
            {
                OutPutFileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
            {
                OutPutFileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            string fileFullname = Path.Combine(OutPutFileDirectory, "Products_" + Guid.NewGuid() + ".xlsx");
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullname, SpreadsheetDocumentType.Workbook))
            {
                CreatePartsForExcel(package, productList);
            }
            return fileFullname;
        }
        private static void CreatePartsForExcel(SpreadsheetDocument package, List<Product> productList)
        {
            SheetData partSheetData = GenerateSheetdataForDetails(productList);

            WorkbookPart workbookPart1 = package.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPartContent(workbookStylesPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartContent(worksheetPart1, partSheetData);
        }
        private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color6);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)3U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
        private static void GenerateWorkbookPartContent(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets1.Append(sheet1);
            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
        }
        private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart1, SheetData partSheetData)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(partSheetData);
            worksheet1.Append(pageMargins1);
            worksheetPart1.Worksheet = worksheet1;
        }
        private static SheetData GenerateSheetdataForDetails(List<Product> productList)
        {
            SheetData sheetData1 = new SheetData();
            sheetData1.Append(CreateHeaderRowForExcel());

            foreach (Product item in productList)
            {
                Row partsRows = GenerateRowForChildPartDetail(item);
                sheetData1.Append(partsRows);
            }
            return sheetData1;
        }
        private static Row GenerateRowForChildPartDetail(Product product)
        {
            Row tRow = new Row();
            tRow.Append(CreateCell(product.ProductId.ToString()));
            tRow.Append(CreateCell(product.Name));
            tRow.Append(CreateCell(product.BrandId.ToString()));
            tRow.Append(CreateCell(product.MfgDate.ToString()));
            tRow.Append(CreateCell(product.ManufactureMonth.ToString()));
            tRow.Append(CreateCell(product.ManufactureYear.ToString()));

            return tRow;
        }
        private static Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.StyleIndex = 1U;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private static Row CreateHeaderRowForExcel()
        {
            Row workRow = new Row();

            workRow.Append(CreateCell("Product Id", 2U));
            workRow.Append(CreateCell("Name", 2U));
            workRow.Append(CreateCell("Brand Id", 2U));
            workRow.Append(CreateCell("Manufacutring Date", 2U));
            workRow.Append(CreateCell("Manufacutring Month", 2U));
            workRow.Append(CreateCell("Manufacutring Year", 2U));
            return workRow;
        }
        private static Cell CreateCell(string text, uint styleIndex)
        {
            Cell cell = new Cell();
            cell.StyleIndex = styleIndex;
            cell.DataType = ResolveCellDataTypeOnValue(text);
            cell.CellValue = new CellValue(text);
            return cell;
        }
        private static EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
        {
            int intVal;
            double doubleVal;
            if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
            {
                return CellValues.Number;
            }
            else
            {
                return CellValues.String;
            }
        }
        #endregion


        #region ImportFromExcel Methods
        private static List<Product> ImportFromExcel()
        {
            List<Product> products = new List<Product>();
            try
            {
                //specify the file name where its actually exist   
                string filepath = @"C:\Users\Karan Khamkar\Desktop\Products_ec79b3ca-0f0d-48cc-af53-62d5b5df34da.xlsx";

                //open the excel using openxml sdk  
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
                {
                    //create the object for workbook part  
                    WorkbookPart wbPart = doc.WorkbookPart;
                    //statement to get the count of the worksheet  
                    int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count();

                    //statement to get the sheet object  
                    Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);

                    //statement to get the worksheet object by using the sheet id  
                    Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;

                    //Note: worksheet has 8 children and the first child[1] = sheetviewdimension,....child[4]=sheetdata  
                    int wkschildno = 4;

                    //statement to get the sheetdata which contains the rows and cell in table  
                    SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);

                    for (int i = 1; i < Rows.Count(); i++)
                    {
                        Row currentrow = (Row)Rows.ChildElements.GetItem(i);

                        Product p = GetProductFromRow(currentrow, wbPart);
                        products.Add(p);
                    }
                }
            }
            catch (Exception Ex)
            {
                _ = Ex.Message;
            }
            return products;
        }
        private static Product GetProductFromRow(Row currentrow, WorkbookPart wbPart)
        {
            Product product = null;
            string productIdFromExcel = ((Cell)currentrow.ChildElements.GetItem(0)).InnerText;
            string productName = GetCellvalue((Cell)currentrow.ChildElements.GetItem(1), wbPart);
            string brandIdFromExcel = ((Cell)currentrow.ChildElements.GetItem(2)).InnerText;
            string manufacturingDateFromExcel = GetCellvalue((Cell)currentrow.ChildElements.GetItem(3), wbPart);
            int productId = int.Parse(productIdFromExcel);
            int brandId = int.Parse(brandIdFromExcel);
            //double dateFromExcel = double.Parse(manufacturingDateFromExcel);
            //var date = DateTime.FromOADate(dateFromExcel);
            DateTime date = DateTime.Parse(manufacturingDateFromExcel);
            product = new Product(productId, productName, brandId, date);
            return product;
        }

        private static string GetCellvalue(Cell currentcell, WorkbookPart wbPart)
        {
            string currentcellvalue = string.Empty;
            if (currentcell.DataType != null)
            {
                if (currentcell.DataType == CellValues.SharedString)
                {
                    int id = -1;

                    if (Int32.TryParse(currentcell.InnerText, out id))
                    {
                        SharedStringItem item = GetSharedStringItemById(wbPart, id);

                        if (item.Text != null)
                        {
                            //code to take the string value  
                            currentcellvalue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            currentcellvalue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            currentcellvalue = item.InnerXml;
                        }
                    }
                }
            }
            return currentcellvalue;
        }
        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        #endregion

        //private static string IsNameValid(string name)
        //{
        //    string result = "-1";
        //    if (name.Length <= 100)
        //    {
        //        result = name;
        //    }
        //    else
        //    {
        //        Console.WriteLine("Product name exceeds the character limit.\nPlease Re-enter name.");
        //        IsNameValid(Console.ReadLine());
        //    }
        //    return result;
        //}

        //private static bool IsValidBrandId(string id, out int brandId)
        //{
        //    bool result = false;
        //    brandId = -1;
        //    int parsedInteger;
        //    bool isValidInteger = int.TryParse(id, out parsedInteger);
        //    if (isValidInteger)
        //    {
        //        var brandList = DataStore.GetAllBrands();

        //        foreach (var item in brandList)
        //        {
        //            if (item.Id == parsedInteger)
        //            {
        //                result = true;
        //                brandId = parsedInteger;
        //                break;
        //            }
        //        } 
        //    }

        //    return result;
        //}
        //private static DateTime IsValidManufactureDate(string date)
        //{
        //    DateTime result;
        //    bool input = DateTime.TryParse(date, out result);
        //    if (input)
        //    {
        //        if (result < DateTime.Now)
        //        {
        //            return result;
        //        }
        //        else
        //        {
        //            Console.WriteLine("Future date not Acceptable.\nPlease Re-enter Date.");
        //            IsValidManufactureDate(Console.ReadLine());
        //        }
        //    }
        //    else
        //    {
        //        Console.WriteLine("Invalid Date.\nPlease Re-enter Date.");
        //        IsValidManufactureDate(Console.ReadLine());
        //    }
        //    return result;
        //}
        #endregion

    }
}
