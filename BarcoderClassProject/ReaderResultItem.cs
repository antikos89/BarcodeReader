using BarcoderClassProject.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BarcoderClassProject
{
    public class ReaderResultItem
    {
        public string PdfFileName { get; set; }

        public int BarcodeNumPage { get; set; }

        public string BarcodeType { get; set; }

        public string BarcodeValue { get; set; }


        public ReaderResultItem()
        {

        }


        public ReaderResultItem(string pdfFileName, int barcodeNumPage, string barcodeType, string barcodeValue)
        {
            PdfFileName = pdfFileName;
            BarcodeNumPage = barcodeNumPage;
            BarcodeType = barcodeType;
            BarcodeValue = barcodeValue;
        }


        public static IEnumerable<ReaderResultItem> ProcessingPdfFile(string pdfFileName, string tempDir, int maxThreads,int scanResolution,int heightPersentageArea)
        {

            #region TimerSet
            Stopwatch timer = new Stopwatch();
            timer.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";
            #endregion
            var pdfPageFileNameShort = Path.GetFileName(pdfFileName);
            var foundBarcodes = new List<ReaderResultItem>();

            var tempFolderForPdf = Path.Combine(tempDir, "PDFProcessingAntonovKA" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempFolderForPdf);
            if (Directory.Exists(tempFolderForPdf))
            {

                var newfileNameForParce = Path.Combine(tempFolderForPdf, Path.GetFileName(pdfFileName));
                File.Copy(pdfFileName, newfileNameForParce);


                ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);
                ConsoleAwesome.WriteLine(String.Format("В работе файл: {0} ", pdfPageFileNameShort), ConsoleColor.DarkGreen);
                ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);


                var barcodeReader = new CustomBarcodeReader(newfileNameForParce, scanResolution, heightPersentageArea);

                var foundBarcodesCount = barcodeReader.BarcodeValuesAndTypes.Count;


                if (foundBarcodesCount > 0)
                {

                    foreach (var bar in barcodeReader.BarcodeValuesAndTypes)
                    {
                        var foundBarcode = new ReaderResultItem();

                        foundBarcode.PdfFileName = pdfPageFileNameShort;
                        foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                        foundBarcode.BarcodeType = bar.Item1;
                        foundBarcode.BarcodeValue = bar.Item2;

                        foundBarcodes.Add(foundBarcode);
                    }


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                    ConsoleAwesome.WriteLine(String.Format("В файле: {0} найдено баркодов: {1}", pdfPageFileNameShort, foundBarcodesCount), ConsoleColor.DarkYellow);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                }
                else
                {

                    var foundBarcode = new ReaderResultItem();

                    foundBarcode.PdfFileName = pdfPageFileNameShort;
                    foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                    foundBarcode.BarcodeType = "";
                    foundBarcode.BarcodeValue = "";

                    foundBarcodes.Add(foundBarcode);


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                    ConsoleAwesome.WriteLine(String.Format("В файле: {0} баркодов не найденно", pdfPageFileNameShort), ConsoleColor.DarkRed);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                }







                #region TimerStop
                timer.Stop();
                totalSeconds = timer.Elapsed.TotalSeconds;
                totalTime += totalSeconds;
                totalSecondsValue = string.Format("{0:0.00}", totalSeconds);
                #endregion


                ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);
                ConsoleAwesome.WriteLine(String.Format("На обработку файла: {0} затрачено = {1} sec.", pdfPageFileNameShort, totalSecondsValue), ConsoleColor.Green);
                ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);



                //Очищаем временную папку
                Directory.Delete(tempFolderForPdf, true);

                return foundBarcodes;

            }
            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pdfFileName">Полное имя PDF файла</param>
        /// <param name="pdfFilestream">MemoryStream PDF файла</param>
        /// <param name="tempDir">Временная папка для сохранения необходимых в обработке файлов</param>
        /// <param name="maxThreads">Количество одновременно обрабатываемых файлов</param>
        /// <param name="scanResolution">Разрешение скана при сохранении распознанного баркода</param>
        /// <param name="heightPersentageArea">Область поиска баркода в % по высоте файла</param>
        /// <param name="saveOriginalPngInTemp">Сохранять ли во временную пакпку обрабатываемые PNG или хранить их в памяти</param>
        /// <returns></returns>
        public static IEnumerable<ReaderResultItem> ProcessingPdfFile(string pdfFileName, MemoryStream pdfFilestream, string tempDir, int maxThreads, int scanResolution, int heightPersentageArea, bool saveOriginalPngInTemp = true)
        {

            #region TimerSet
            Stopwatch timer = new Stopwatch();
            timer.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";
            #endregion
            var pdfPageFileNameShort = Path.GetFileName(pdfFileName);
            var foundBarcodes = new List<ReaderResultItem>();

            var tempFolderForPdf = Path.Combine(tempDir, "PDFProcessingAntonovKA" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempFolderForPdf);
            if (Directory.Exists(tempFolderForPdf))
            {

                //var newfileNameForParce = Path.Combine(tempFolderForPdf, Path.GetFileName(pdfFileName));
                //File.Copy(pdfFileName, newfileNameForParce);


                ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);
                ConsoleAwesome.WriteLine(String.Format("В работе файл: {0} ", pdfPageFileNameShort), ConsoleColor.DarkGreen);
                ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);


                var barcodeReader = new CustomBarcodeReader(pdfFileName, pdfFilestream, tempFolderForPdf, scanResolution, heightPersentageArea, saveOriginalPngInTemp);

                var foundBarcodesCount = barcodeReader.BarcodeValuesAndTypes.Count;


                if (foundBarcodesCount > 0)
                {

                    foreach (var bar in barcodeReader.BarcodeValuesAndTypes)
                    {
                        var foundBarcode = new ReaderResultItem();

                        foundBarcode.PdfFileName = pdfPageFileNameShort;
                        foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                        foundBarcode.BarcodeType = bar.Item1;
                        foundBarcode.BarcodeValue = bar.Item2;

                        foundBarcodes.Add(foundBarcode);
                    }


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                    ConsoleAwesome.WriteLine(String.Format("В файле: {0} найдено баркодов: {1}", pdfPageFileNameShort, foundBarcodesCount), ConsoleColor.DarkYellow);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                }
                else
                {

                    var foundBarcode = new ReaderResultItem();

                    foundBarcode.PdfFileName = pdfPageFileNameShort;
                    foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                    foundBarcode.BarcodeType = "";
                    foundBarcode.BarcodeValue = "";

                    foundBarcodes.Add(foundBarcode);


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                    ConsoleAwesome.WriteLine(String.Format("В файле: {0} баркодов не найденно", pdfPageFileNameShort), ConsoleColor.DarkRed);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                }







                #region TimerStop
                timer.Stop();
                totalSeconds = timer.Elapsed.TotalSeconds;
                totalTime += totalSeconds;
                totalSecondsValue = string.Format("{0:0.00}", totalSeconds);
                #endregion


                ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);
                ConsoleAwesome.WriteLine(String.Format("На обработку файла: {0} затрачено = {1} sec.", pdfPageFileNameShort, totalSecondsValue), ConsoleColor.Green);
                ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);



                //Очищаем временную папку
                Directory.Delete(tempFolderForPdf, true);

                return foundBarcodes;

            }
            return null;
        }



        public static IEnumerable<ReaderResultItem> ProcessingPdfFile(string pdfFileName,MemoryStream pdfFilestream, string tempDir, int maxThreads, int scanResolution, int heightPersentageArea)
        {

            #region TimerSet
            Stopwatch timer = new Stopwatch();
            timer.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";
            #endregion
            
            var pdfPageFileNameShort = Path.GetFileName(pdfFileName);
            var foundBarcodes = new List<ReaderResultItem>();

            var tempFolderForPdf = Path.Combine(tempDir, "PDFProcessingAntonovKA" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempFolderForPdf);
            if (Directory.Exists(tempFolderForPdf))
            {

                //var newfileNameForParce = Path.Combine(tempFolderForPdf, Path.GetFileName(pdfFileName));
                //File.Copy(pdfFileName, newfileNameForParce);


                ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);
                ConsoleAwesome.WriteLine(String.Format("В работе файл: {0} ", pdfPageFileNameShort), ConsoleColor.DarkGreen);
                ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);


                var barcodeReader = new CustomBarcodeReader(pdfFileName, pdfFilestream, tempFolderForPdf, scanResolution, heightPersentageArea);

                var foundBarcodesCount = barcodeReader.BarcodeValuesAndTypes.Count;


                if (foundBarcodesCount > 0)
                {

                    foreach (var bar in barcodeReader.BarcodeValuesAndTypes)
                    {
                        var foundBarcode = new ReaderResultItem();

                        foundBarcode.PdfFileName = pdfPageFileNameShort;
                        foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                        foundBarcode.BarcodeType = bar.Item1;
                        foundBarcode.BarcodeValue = bar.Item2;

                        foundBarcodes.Add(foundBarcode);
                    }


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                    ConsoleAwesome.WriteLine(String.Format("В файле: {0} найдено баркодов: {1}", pdfPageFileNameShort, foundBarcodesCount), ConsoleColor.DarkYellow);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                }
                else
                {

                    var foundBarcode = new ReaderResultItem();

                    foundBarcode.PdfFileName = pdfPageFileNameShort;
                    foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                    foundBarcode.BarcodeType = "";
                    foundBarcode.BarcodeValue = "";

                    foundBarcodes.Add(foundBarcode);


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                    ConsoleAwesome.WriteLine(String.Format("В файле: {0} баркодов не найденно", pdfPageFileNameShort), ConsoleColor.DarkRed);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                }







                #region TimerStop
                timer.Stop();
                totalSeconds = timer.Elapsed.TotalSeconds;
                totalTime += totalSeconds;
                totalSecondsValue = string.Format("{0:0.00}", totalSeconds);
                #endregion


                ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);
                ConsoleAwesome.WriteLine(String.Format("На обработку файла: {0} затрачено = {1} sec.", pdfPageFileNameShort, totalSecondsValue), ConsoleColor.Green);
                ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);



                //Очищаем временную папку
                Directory.Delete(tempFolderForPdf, true);

                return foundBarcodes;

            }
            return null;
        }


        public async static Task<IEnumerable<ReaderResultItem>> ProcessingPdfFileTask(string pdfFileName, string tempDir, int maxThreads, int scanResolution, int heightPersentageArea)
        {

            #region TimerSet
            Stopwatch timer = new Stopwatch();
            timer.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";
            #endregion
            var pdfPageFileNameShort = Path.GetFileName(pdfFileName);
            var foundBarcodes = new List<ReaderResultItem>();

            var resTask = Task.Run(
            () => 
            {

                var tempFolderForPdf = Path.Combine(tempDir, "PDFProcessingAntonovKA" + Path.GetRandomFileName());
                Directory.CreateDirectory(tempFolderForPdf);
                if (Directory.Exists(tempFolderForPdf))
                {

                    var newfileNameForParce = Path.Combine(tempFolderForPdf, Path.GetFileName(pdfFileName));
                    File.Copy(pdfFileName, newfileNameForParce);


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);
                    ConsoleAwesome.WriteLine(String.Format("В работе файл: {0} ", pdfPageFileNameShort), ConsoleColor.DarkGreen);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);


                    var barcodeReader = new CustomBarcodeReader(newfileNameForParce, scanResolution, heightPersentageArea);

                    var foundBarcodesCount = barcodeReader.BarcodeValuesAndTypes.Count;


                    if (foundBarcodesCount > 0)
                    {

                        foreach (var bar in barcodeReader.BarcodeValuesAndTypes)
                        {
                            var foundBarcode = new ReaderResultItem();

                            foundBarcode.PdfFileName = pdfPageFileNameShort;
                            foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                            foundBarcode.BarcodeType = bar.Item1;
                            foundBarcode.BarcodeValue = bar.Item2;

                            foundBarcodes.Add(foundBarcode);
                        }


                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                        ConsoleAwesome.WriteLine(String.Format("В файле: {0} найдено баркодов: {1}", pdfPageFileNameShort, foundBarcodesCount), ConsoleColor.DarkYellow);
                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                    }
                    else
                    {

                        var foundBarcode = new ReaderResultItem();

                        foundBarcode.PdfFileName = pdfPageFileNameShort;
                        foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                        foundBarcode.BarcodeType = "";
                        foundBarcode.BarcodeValue = "";

                        foundBarcodes.Add(foundBarcode);


                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                        ConsoleAwesome.WriteLine(String.Format("В файле: {0} баркодов не найденно", pdfPageFileNameShort), ConsoleColor.DarkRed);
                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                    }







                    #region TimerStop
                    timer.Stop();
                    totalSeconds = timer.Elapsed.TotalSeconds;
                    totalTime += totalSeconds;
                    totalSecondsValue = string.Format("{0:0.00}", totalSeconds);
                    #endregion


                    ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);
                    ConsoleAwesome.WriteLine(String.Format("На обработку файла: {0} затрачено = {1} sec.", pdfPageFileNameShort, totalSecondsValue), ConsoleColor.Green);
                    ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);



                    //Очищаем временную папку
                    Directory.Delete(tempFolderForPdf, true);

                    

                }
                


            }).ConfigureAwait(false);

            await resTask;
            return foundBarcodes;


        }


        public async static Task<IEnumerable<ReaderResultItem>> ProcessingPdfFileTask(string pdfFileName, MemoryStream pdfFilestream, string tempDir, int maxThreads, int scanResolution, int heightPersentageArea,bool saveOriginalPngInTemp=true)
        {

            #region TimerSet
            Stopwatch timer = new Stopwatch();
            timer.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";
            #endregion
            var pdfPageFileNameShort = Path.GetFileName(pdfFileName);
            var foundBarcodes = new List<ReaderResultItem>();

            var resTask = Task.Run(
            () =>
            {

                var tempFolderForPdf = Path.Combine(tempDir, "PDFProcessingAntonovKA" + Path.GetRandomFileName());
                Directory.CreateDirectory(tempFolderForPdf);
                if (Directory.Exists(tempFolderForPdf))
                {

                    //var newfileNameForParce = Path.Combine(tempFolderForPdf, Path.GetFileName(pdfFileName));
                    //File.Copy(pdfFileName, newfileNameForParce);


                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);
                    ConsoleAwesome.WriteLine(String.Format("В работе файл: {0} ", pdfPageFileNameShort), ConsoleColor.DarkGreen);
                    ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkGreen);


                    var barcodeReader = new CustomBarcodeReader(pdfFileName, pdfFilestream, tempFolderForPdf, scanResolution, heightPersentageArea, saveOriginalPngInTemp);

                    var foundBarcodesCount = barcodeReader.BarcodeValuesAndTypes.Count;


                    if (foundBarcodesCount > 0)
                    {

                        foreach (var bar in barcodeReader.BarcodeValuesAndTypes)
                        {
                            var foundBarcode = new ReaderResultItem();

                            foundBarcode.PdfFileName = pdfPageFileNameShort;
                            foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                            foundBarcode.BarcodeType = bar.Item1;
                            foundBarcode.BarcodeValue = bar.Item2;

                            foundBarcodes.Add(foundBarcode);
                        }


                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                        ConsoleAwesome.WriteLine(String.Format("В файле: {0} найдено баркодов: {1}", pdfPageFileNameShort, foundBarcodesCount), ConsoleColor.DarkYellow);
                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkYellow);
                    }
                    else
                    {

                        var foundBarcode = new ReaderResultItem();

                        foundBarcode.PdfFileName = pdfPageFileNameShort;
                        foundBarcode.BarcodeNumPage = barcodeReader.BarcodeNumPage;
                        foundBarcode.BarcodeType = "";
                        foundBarcode.BarcodeValue = "";

                        foundBarcodes.Add(foundBarcode);


                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                        ConsoleAwesome.WriteLine(String.Format("В файле: {0} баркодов не найденно", pdfPageFileNameShort), ConsoleColor.DarkRed);
                        ConsoleAwesome.WriteLine("-----------------------------------------------------------", ConsoleColor.DarkRed);
                    }







                    #region TimerStop
                    timer.Stop();
                    totalSeconds = timer.Elapsed.TotalSeconds;
                    totalTime += totalSeconds;
                    totalSecondsValue = string.Format("{0:0.00}", totalSeconds);
                    #endregion


                    ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);
                    ConsoleAwesome.WriteLine(String.Format("На обработку файла: {0} затрачено = {1} sec.", pdfPageFileNameShort, totalSecondsValue), ConsoleColor.Green);
                    ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Green);



                    //Очищаем временную папку
                    Directory.Delete(tempFolderForPdf, true);



                }



            }).ConfigureAwait(false);

            await resTask;
            return foundBarcodes;


        }


        public static void CreateExcell(string outputFilePath, string outputFileName, IEnumerable<ReaderResultItem> items)
        {

            var fullPath = System.IO.Path.Combine(outputFilePath, string.Format(@"{0}_{1}.xlsx", outputFileName, DateTime.Now.ToString("MM_dd_yyyy HH_mm_ss")));
            System.IO.FileInfo newFile = new System.IO.FileInfo(fullPath);

            if (newFile.Exists)
            {
                newFile.Delete();

                newFile = new System.IO.FileInfo(fullPath);
            }



            using (ExcelPackage excelPackage = new ExcelPackage(newFile))
            {

                if (items.Count() > 0)
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("BarcodeValues");


                    worksheet.Cells[1, 1].Value = "ИМЯ ФАЙЛА";
                    worksheet.Cells[1, 2].Value = "№ СТРАНИЦЫ PDF";
                    worksheet.Cells[1, 3].Value = "ТИП БАРКОДА";
                    worksheet.Cells[1, 4].Value = "№ БАРКОДА";



                    var rowIdx = 2;



                    foreach (var item in items)
                    {
                        worksheet.Cells[rowIdx, 1].Value = item.PdfFileName;
                        worksheet.Cells[rowIdx, 2].Value = item.BarcodeNumPage;
                        worksheet.Cells[rowIdx, 3].Value = item.BarcodeType;
                        worksheet.Cells[rowIdx, 4].Value = item.BarcodeValue;

                        rowIdx++;
                    }

                    SetTableStyle(worksheet);
                }


                excelPackage.Save();
            }

        }


        public static void SetTableStyle(ExcelWorksheet table)
        {
            var colCount = table.Dimension.End.Column;
            table.Row(1).Height = 32;

            for (int i = 1; i <= colCount; i++)
            {
                table.Column(i).AutoFit();
                table.Column(i).Style.WrapText = true;
            }

            using (var range = table.Cells[1, 1, 1, colCount])
            {

                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(189, 215, 238));
                range.Style.Font.Color.SetColor(Color.Black);
                range.Style.Font.Size = 11;
                range.Style.Font.Bold = true;
                range.Style.Font.Name = "Calibri";
                range.Style.Border.Left.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Top.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Right.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                range.Style.Border.Top.Color.SetColor(Color.Black);
                range.Style.Border.Bottom.Color.SetColor(Color.Black);
                range.Style.Border.Left.Color.SetColor(Color.Black);
                range.Style.Border.Right.Color.SetColor(Color.Black);
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            var rowCount = table.Dimension.End.Row;
            using (var range = table.Cells[2, 1, rowCount, colCount])
            {
                range.Style.Font.Color.SetColor(Color.Black);
                range.Style.Font.Size = 11;
                range.Style.Font.Name = "Calibri";
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Top.Color.SetColor(Color.Black);
                range.Style.Border.Bottom.Color.SetColor(Color.Black);
                range.Style.Border.Left.Color.SetColor(Color.Black);
                range.Style.Border.Right.Color.SetColor(Color.Black);
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            using (var range = table.Cells[1, 1, rowCount, colCount])
            {
                range.AutoFilter = true;
            }
        }


    }
}
