using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;

namespace ccsxmltoxlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            const String ROOT_ELEMENT = "Workbook";
            const String WORKSHEET_ELEMENT = "Worksheet";
            const String TABLE_ELEMENT = "Table";
            const String ROW_ELEMENT = "Row";
            const String CELL_ELEMENT = "Cell";
            const String CUSTOMER_COLUMN = "Customer";
            const String DICTIONARY_COLUMN = "Dictionary";
            const String VARIABLE_COLUMN = "Variable";
            const String VALUE_COLUMN = "Value";

            if (args.Count() == 0) {
                Console.WriteLine("> xls file is missing.");
            } else if (args.Count() == 1)
            {
                Console.WriteLine("> Input file is missing.");
            } else
            {
                // Read xsl file
                String xslfilename = args[0];
                Console.WriteLine("xls file is " + xslfilename + ".");
                var xslt = new XslCompiledTransform();
                xslt.Load(xslfilename);

                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                dt.Columns.Add(CUSTOMER_COLUMN, Type.GetType("System.String"));
                dt.Columns.Add(DICTIONARY_COLUMN, Type.GetType("System.String"));
                dt.Columns.Add(VARIABLE_COLUMN, Type.GetType("System.String"));
                dt.Columns.Add(VALUE_COLUMN, Type.GetType("System.String"));
                ds.Tables.Add(dt);

                var xlsxWorkbook = new XLWorkbook();
                var xlsxWorksheet = xlsxWorkbook.Worksheets.Add("Parameters");
                
                for (int i = 1; i < args.Length; ++i)
                {
                    Console.WriteLine("Input file is " + args[i] + ".");

                    var output = new StringBuilder();
                    var settings = new XmlWriterSettings();
                    settings.OmitXmlDeclaration = true;
                    settings.Indent = true;
                    String customer = Path.GetFileNameWithoutExtension(args[i]);
                    String modifiedXMLFilename = Path.GetFileNameWithoutExtension(args[i]) + "_modified.xml";

                    using (var writer = XmlWriter.Create(output, settings))
                    {
                        xslt.Transform(args[i], writer);
                    }


                    XDocument xmlDoc = XDocument.Parse(output.ToString());
                    XElement workbook = xmlDoc.Element(ROOT_ELEMENT);
                    XElement worksheet = workbook.Element(WORKSHEET_ELEMENT);
                    XElement table = worksheet.Element(TABLE_ELEMENT);
                    var rows = table.Elements(ROW_ELEMENT);
                    

                    foreach (XElement row in rows)
                    {
                        var cells = row.Elements(CELL_ELEMENT).ToArray<XElement>();
                        DataRow dtRow = dt.NewRow();
                        dtRow[CUSTOMER_COLUMN] = customer;
                        dtRow[DICTIONARY_COLUMN] = cells[0].Value;
                        dtRow[VARIABLE_COLUMN] = cells[1].Value;
                        dtRow[VALUE_COLUMN] = cells[2].Value;
                        dt.Rows.Add(dtRow);

                        //Console.WriteLine(dtRow[CUSTOMER_COLUMN] + "," + dtRow[DICTIONARY_COLUMN] + "," + dtRow[VARIABLE_COLUMN] + "," + dtRow[VALUE_COLUMN]);                       
                    }
                }


                xlsxWorksheet.Cell(1, 1).Value = CUSTOMER_COLUMN;
                xlsxWorksheet.Cell(1, 2).Value = DICTIONARY_COLUMN;
                xlsxWorksheet.Cell(1, 3).Value = VARIABLE_COLUMN;
                xlsxWorksheet.Cell(1, 4).Value = VALUE_COLUMN;
                xlsxWorksheet.Cell(1, 5).Value = "Note";

                int rowIndex = 2;
                DataRow[] dtRows = dt.Select(null, DICTIONARY_COLUMN + ", " + VARIABLE_COLUMN + " ASC");

                foreach (var dtRow in dtRows)
                {
                    xlsxWorksheet.Cell(rowIndex, 1).Value = dtRow[CUSTOMER_COLUMN];
                    xlsxWorksheet.Cell(rowIndex, 2).Value = dtRow[DICTIONARY_COLUMN];
                    xlsxWorksheet.Cell(rowIndex, 3).Value = dtRow[VARIABLE_COLUMN];
                    xlsxWorksheet.Cell(rowIndex, 4).Value = dtRow[VALUE_COLUMN];
                    rowIndex++;
                }

                try
                {
                    xlsxWorkbook.SaveAs("CCSParamComp.xlsx");
                } catch (System.IO.IOException ex)
                {
                    Console.WriteLine("Excelファイルの保存に失敗しました。");
                }
            }
        }
    }
}
