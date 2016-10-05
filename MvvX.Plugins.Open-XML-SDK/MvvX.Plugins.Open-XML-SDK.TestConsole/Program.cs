﻿using System;
using System.IO;
using MvvX.Plugins.OpenXMLSDK.Word;
using System.Collections.Generic;
using MvvX.Plugins.OpenXMLSDK.Word.Tables.Models;
using MvvX.Plugins.OpenXMLSDK.Word.Tables;
using MvvX.Plugins.OpenXMLSDK.Word;
using MvvX.Plugins.OpenXMLSDK.Word.Paragraphs;
using System.Diagnostics;
using MvvX.Plugins.OpenXMLSDK.Word.Models;
using MvvmCross.Platform;
using MvvX.Plugins.OpenXMLSDK.Platform.Word;
using MvvX.Plugins.OpenXMLSDK.Platform.Excel;
using MvvX.Plugins.OpenXMLSDK.Excel;

namespace MvvX.Plugins.OpenXMLSDK.TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            //var resourceName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Global.docx");

            //if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results")))
            //    Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results"));

            //string finalFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "Results", "FinalDoc_Test_OrientationParagraph-" + DateTime.Now.ToFileTime() + ".docx");

            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample.xlsx")))
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample.xlsx"));

            FileStream fs = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample.xlsx"), FileMode.OpenOrCreate, FileAccess.ReadWrite);

            using (IExcelManager excel = new ExcelManager())
            {

                excel.CreateDoc(fs);

                // add a new worksheet to the empty workbook
                excel.CreateWorksheet("First Worksheet");

                //add the headers
                for (int i = 1 ; i < 5; i++)
                {
                    excel.AddCell(1, i, "Column" + i);
                }

                //add some values
                excel.AddCell("A2", 4564);
                excel.AddCell("A3", "Cell A3");
                excel.AddCell("C5", "Hi !");
                excel.AddCell("B2", 45);
                excel.AddCell("D2", 787);

                excel.Save();
            }

            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample2.xlsx")))
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample2.xlsx"));

            FileStream instream = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "template.xlsx"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
            FileStream stream = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App_Data", "sample2.xlsx"), FileMode.OpenOrCreate, FileAccess.ReadWrite);

            using (IExcelManager excel = new ExcelManager())
            {
                excel.CreateDocFromTemplate(stream, instream);

                //give current worksheet name 
                excel.WorksheetName = "First Worksheet";
                //or number
                excel.WorksheetNumber = 1;

                //add some values
                excel.AddCell("A2", 1111);
                excel.AddCell("A3", "000");
                excel.AddCell("C5", "Test template");

                excel.Save(stream);
                instream.Close();
            }




            //    using (IWordManager word = new WordManager())
            //    {
            //        // TODO for debug : use your test file :
            //        word.OpenDocFromTemplate(resourceName, finalFilePath, true);

            //    //    word.SaveDoc();
            //    //    word.CloseDoc();
            //    //}
            //    // Insertion de texte dans un bookmark
            //    // wordManager.SetTextOnBookmark("Insert_Documents", "Hi !");

            //    // Insertion d'une table dans un bookmark
            //    // Propriété du Tableau
            //    var tableProperty = new TablePropertiesModel()
            //        {
            //            TableBorders = new TableBordersModel()
            //            {
            //                TopBorder = new TableBorderModel() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds },
            //                LeftBorder = new TableBorderModel() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
            //                RightBorder = new TableBorderModel() { Color = "CCCCCC", Size = 20, BorderValue = BorderValues.Birds },
            //                BottomBorder = new TableBorderModel() { Color = "F7941F", Size = 40, BorderValue = BorderValues.Birds }
            //            },
            //            TableWidth = new TableWidthModel()
            //            {
            //                Width = "5000",
            //                Type = TableWidthUnitValues.Pct
            //            }
            //        };
            //        // Lignes du premier tableau pour les constats checked
            //        var lines = new List<ITableRow>();

            //        for (int i = 0; i < 3; i++)
            //        {
            //            var borderTopIsOK = new TableBorderModel();
            //            if (i != 0)
            //                borderTopIsOK.BorderValue = BorderValues.Nil;

            //            // Première ligne
            //            var text = word.CreateRunForText("Header Number : " + i,
            //                    new RunPropertiesModel()
            //                    {
            //                        Bold = true,
            //                        FontSize = "24",
            //                        RunFonts = new RunFontsModel()
            //                        {
            //                            Ascii = "Courier New",
            //                            HighAnsi = "Courier New",
            //                            EastAsia = "Courier New",
            //                            ComplexScript = "Courier New"
            //                        }
            //                    });

            //            var cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(text, new TableCellPropertiesModel() {
            //                    Gridspan = new GridSpanModel() { Val = 2 },
            //                    Shading = new ShadingModel()
            //                    {
            //                        Fill = "F7941F"
            //                    },
            //                    TableCellWidth = new TableCellWidthModel()
            //                    {
            //                        Width = "8862"
            //                    },
            //                    TableCellBorders = new TableCellBordersModel()
            //                    {
            //                        TopBorder = borderTopIsOK
            //                    }
            //                }),
            //                word.CreateTableCell(word.CreateRun(), new TableCellPropertiesModel() { 
            //                            TableCellWidth = new TableCellWidthModel()
            //                            {
            //                                Width = "246"
            //                            },
            //                            Shading = new ShadingModel()
            //                            {
            //                                Fill = "F7941F"
            //                            },
            //                            TableCellBorders = new TableCellBordersModel() {
            //                                        TopBorder = borderTopIsOK
            //                            }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules, new TableRowPropertiesModel()
            //            {
            //                TableRowHeight = new TableRowHeightModel()
            //                {
            //                    Val = 380
            //                }
            //            }));

            //            // Deuxième ligne
            //            text = word.CreateRunForText("Comments", new RunPropertiesModel() { Bold = true });
            //            cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(word.CreateImage(@"c:\temp\Tulips.jpg", new Drawing.Pictures.Model.PictureModel() {
            //                    ImagePartType   = Packaging.ImagePartType.Jpeg,
            //                    MaxHeight = 10,
            //                    MaxWidth = 500
            //                }), new TableCellPropertiesModel() {
            //                            TableCellWidth = new TableCellWidthModel()
            //                            {
            //                                Width = "4890"
            //                            }
            //                }),
            //                word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
            //                            Fusion = true,
            //                            TableCellWidth = new TableCellWidthModel()
            //                            {
            //                                Width = "4218"
            //                            },
            //                            Gridspan = new GridSpanModel() { Val = 2 }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules));

            //            // Troisième ligne
            //            text = word.CreateRunForText("Texte du Constat Number : " + i, new RunPropertiesModel());
            //            cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(text, new TableCellPropertiesModel() {
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4890"
            //                                        },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            BottomBorder = new TableBorderModel() {
            //                                                Color = "FF0019"
            //                                                }
            //                                        }
            //                }),
            //                word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
            //                                        Fusion = true,
            //                                        FusionChild = true,
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4218"
            //                                        },
            //                                        Gridspan = new GridSpanModel() { Val = 2 },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            BottomBorder = new TableBorderModel() {
            //                                                Color = "FF0019" }
            //                                            }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules));

            //            // Quatrième ligne
            //            text = word.CreateRunForText("Risques", new RunPropertiesModel() { Bold = true });
            //            cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(text, new TableCellPropertiesModel() {
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4890"
            //                                        },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            TopBorder = new TableBorderModel()
            //                                            {
            //                                                Color = "00FF19"
            //                                            }
            //                                        }
            //                }),
            //                word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
            //                                        Fusion = true,
            //                                        FusionChild = true,
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4218"
            //                                        },
            //                                        Gridspan = new GridSpanModel() { Val = 2 },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            TopBorder = new TableBorderModel()
            //                                            {
            //                                                Color = "00FF19"
            //                                            }
            //                                        }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules));

            //            // Cinquième ligne
            //            text = word.CreateRunForText("Texte du Risque Number : " + i, new RunPropertiesModel());
            //            cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(text, new TableCellPropertiesModel() {
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4890"
            //                                        },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            TopBorder = new TableBorderModel() {
            //                                                BorderValue = BorderValues.Nil }
            //                                        }
            //                }),
            //                word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
            //                                        Fusion = true,
            //                                        FusionChild = true,
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4218"
            //                                        },
            //                                        Gridspan = new GridSpanModel() { Val = 2 },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            TopBorder = new TableBorderModel() {
            //                                                BorderValue = BorderValues.Nil }
            //                                            }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules));

            //            // Sixième ligne
            //            text = word.CreateRunForText("Recommandations", new RunPropertiesModel() { Bold = true });
            //            cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(text, new TableCellPropertiesModel() {
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4890"
            //                                        },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            BottomBorder = new TableBorderModel() {
            //                                                BorderValue = BorderValues.Nil }
            //                                            }
            //                }),
            //                word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
            //                                        Fusion = true,
            //                                        FusionChild = true,
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4218"
            //                                        },
            //                                        Gridspan = new GridSpanModel() { Val = 2 },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            BottomBorder = new TableBorderModel() {
            //                                                BorderValue = BorderValues.Nil }
            //                                            }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules));

            //            var borderBottomIsOK = new TableBorderModel() {
            //                BorderValue = BorderValues.Nil,
            //                Color = "FF0019"
            //            };

            //            if (i == 2)
            //                borderBottomIsOK.BorderValue = BorderValues.Single;

            //            // Septième ligne
            //            text = word.CreateRunForText("Texte de la Recommandation Number : " + i, new RunPropertiesModel());
            //            cellules = new List<ITableCell>()
            //            {
            //                word.CreateTableCell(text, new TableCellPropertiesModel() {
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4890"
            //                                        },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            TopBorder = new TableBorderModel() { BorderValue = BorderValues.Nil },
            //                                            BottomBorder = borderBottomIsOK }
            //                }),
            //                word.CreateTableMergeCell(word.CreateRun(), new TableCellPropertiesModel() {
            //                                        Fusion = true,
            //                                        FusionChild = true,
            //                                        TableCellWidth = new TableCellWidthModel()
            //                                        {
            //                                            Width = "4218"
            //                                        },
            //                                        Gridspan = new GridSpanModel() { Val = 2 },
            //                                        TableCellBorders = new TableCellBordersModel() {
            //                                            TopBorder = new TableBorderModel() {
            //                                                BorderValue = BorderValues.Nil },
            //                                            BottomBorder = borderBottomIsOK }
            //                })
            //            };
            //            lines.Add(word.CreateTableRow(cellules));
            //        }

            //        IList<IParagraph> tables = new List<IParagraph>();
            //        tables.Add(word.CreateParagraphForRun(word.CreateRunForTable(word.CreateTable(lines, tableProperty))));

            //        // Lignes du deuxième tableau pour les constats unchecked
            //        //lines = new List<TableRow>();

            //        if (tables.Count > 0)
            //            word.SetParagraphsOnBookmark("Insert_Documents", tables);

            //        word.SaveDoc();
            //        word.CloseDoc();
            //    }

            //    Process.Start(finalFilePath);
        }
    }
}