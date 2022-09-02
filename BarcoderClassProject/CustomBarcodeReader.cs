using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;

namespace BarcoderClassProject
{
    public class CustomBarcodeReader
    {

        public string PdfFileName { get; set; }

        private List<string> PngFilesFromPdf { get; }

        private Dictionary<int, MemoryStream> PngStreamsFromPdf { get; }

        public int BarcodeNumPage { get; set; }

        public List<Tuple<string, string>> BarcodeValuesAndTypes { get; set; }


        public CustomBarcodeReader(string pdfFileName, int scanResolution = 300, int heightPersentageArea = 10)
        {
            PdfFileName = pdfFileName;
            PngFilesFromPdf = ConvertPdfToPng(pdfFileName, scanResolution);
            BarcodeValuesAndTypes = new List<Tuple<string, string>>();

            //Получаем значения баркодов
            //GetBarcodeValuesAndType3lib(PngFilesFromPdf, heightPersentageArea);
            GetBarcodeValuesAndType4libs(PngFilesFromPdf, heightPersentageArea);
        }


        public CustomBarcodeReader(string pdfFileName, MemoryStream pdfFileStream, string tempDir, int scanResolution = 300, int heightPersentageArea = 10, bool saveOriginalPNGinTemp = true)
        {
            PdfFileName = pdfFileName;



            BarcodeValuesAndTypes = new List<Tuple<string, string>>();
            //
            if (saveOriginalPNGinTemp)
            {
                PngFilesFromPdf = ConvertPdfToPng(pdfFileName, pdfFileStream, tempDir, scanResolution);
                //Получаем значения баркодов
                //GetBarcodeValuesAndType3lib(PngFilesFromPdf, heightPersentageArea);
                GetBarcodeValuesAndType4libs(PngFilesFromPdf, heightPersentageArea);
            }
            else
            {

                PngStreamsFromPdf = ConvertPdfToPngStream(pdfFileName, pdfFileStream, tempDir, scanResolution);
                GetBarcodeValuesAndType4libs(PngStreamsFromPdf, tempDir, heightPersentageArea);
            }



        }


        private List<string> ConvertPdfToPng(string originPath, int resolution = 300)
        {
            List<string> result = new List<string>();
            var directoryName = Path.GetDirectoryName(originPath);
            //var fileNameWithoutExst = Path.GetFileNameWithoutExtension(originPath);

            var listOfPages = SplitBySinglePagesMethod(directoryName, originPath, "_page");

            foreach (var page in listOfPages)
            {
                Spire.Pdf.PdfDocument PdfDocumentSpire = new Spire.Pdf.PdfDocument();
                PdfDocumentSpire.LoadFromFile(page);


                var outputPath = Path.Combine(directoryName, string.Concat(Path.GetFileNameWithoutExtension(page), ".png"));
                var image = PdfDocumentSpire.SaveAsImage(0, Spire.Pdf.Graphics.PdfImageType.Bitmap, resolution, resolution);
                image.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png);
                result.Add(outputPath);

            }



            return result;
        }

        private List<string> ConvertPdfToPng(string originFileName, MemoryStream originFileStream, string outPutDir, int resolution = 300)
        {
            List<string> result = new List<string>();

            //var fileNameWithoutExst = Path.GetFileNameWithoutExtension(originPath);

            var listOfPages = SplitBySinglePagesMethod(outPutDir, originFileName, originFileStream, "_page");

            foreach (var page in listOfPages)
            {
                Spire.Pdf.PdfDocument PdfDocumentSpire = new Spire.Pdf.PdfDocument();
                PdfDocumentSpire.LoadFromBytes(page.Value);


                var outputPath = Path.Combine(outPutDir, string.Concat(Path.GetFileNameWithoutExtension(originFileName), "_page", page.Key, ".png"));
                var image = PdfDocumentSpire.SaveAsImage(0, Spire.Pdf.Graphics.PdfImageType.Bitmap, resolution, resolution);

                image.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png);

                result.Add(outputPath);

            }



            return result;
        }



        private Dictionary<int, MemoryStream> ConvertPdfToPngStream(string originFileName, MemoryStream originFileStream, string outPutDir, int resolution = 300)
        {
            Dictionary<int, MemoryStream> result = new Dictionary<int, MemoryStream>();

            var listOfPages = SplitBySinglePagesMethod(outPutDir, originFileName, originFileStream, "_page");

            foreach (var page in listOfPages)
            {
                Spire.Pdf.PdfDocument PdfDocumentSpire = new Spire.Pdf.PdfDocument();
                PdfDocumentSpire.LoadFromBytes(page.Value);


                var stream = new MemoryStream();
                //var outputPath = Path.Combine(directoryName, string.Concat(Path.GetFileNameWithoutExtension(page), ".png"));
                var image = PdfDocumentSpire.SaveAsImage(0, Spire.Pdf.Graphics.PdfImageType.Bitmap, resolution, resolution);
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                result.Add(page.Key, stream);

            }

            return result;
        }

        /// <summary>
        /// Получает значения баркодов с первой успешно распознанной страницы
        /// </summary>
        /// <param name="pngfiles"></param>
        /// <param name="heightPersentageArea">Область поиска по высоте листа в процентах</param>
        private void GetBarcodeValuesAndType3lib(List<string> pngfiles, int heightPersentageArea)
        {
            foreach (var pngFile in pngfiles)
            {
                //Stream stream = new MemoryStream(File.ReadAllBytes(pngFile));
                var bitmap = new Bitmap(pngFile);

                using (Aspose.BarCode.BarCodeRecognition.BarCodeReader reader = new Aspose.BarCode.BarCodeRecognition.BarCodeReader(bitmap))
                {

                    if (bitmap.Height > 500)
                    {

                        var findArea = new Rectangle(0, 0, bitmap.Width, bitmap.Height / 100 * heightPersentageArea);
                        //var outputFile = Path.Combine(Path.GetDirectoryName(pngFile), String.Concat(Path.GetRandomFileName(), @".png"));
                        //CropImageByRect(pngFile, findArea, outputFile);
                        reader.SetBarCodeImage(bitmap, findArea);
                    }


                    var readerRes = reader.ReadBarCodes();

                    foreach (Aspose.BarCode.BarCodeRecognition.BarCodeResult result in readerRes)
                    {

                        int pdfPageNum = GetNumPageFromPath(pngFile);
                        //Console.WriteLine("Symbology Type: " + result.CodeType);
                        //Console.WriteLine("CodeText: " + result.CodeText);

                        var asposeDirtyRes = result.CodeText;
                        var matchWatermarkRes = Regex.Match(asposeDirtyRes, @"(.{1,})(Recognized)");
                        var resultValue = "";
                        //Если нашли Watermark, то будем получать результат совместно из двух библиотек
                        if (matchWatermarkRes.Success)
                        {
                            var asposeBarcodePart = matchWatermarkRes.Groups[1].Value;

                            //область где найден баркод
                            var findedBarcodeRegion = result.Region;
                            var regionOfBarcode = new Rectangle(findedBarcodeRegion.Rectangle.X - 4, findedBarcodeRegion.Rectangle.Y - 4, findedBarcodeRegion.Rectangle.Width + 8, findedBarcodeRegion.Rectangle.Height + 8);
                            var barcodePath = Path.Combine(Path.GetDirectoryName(pngFile), String.Concat(Path.GetRandomFileName(), @".png"));

                            //Сохраняем только сам баркод под распознование BarcodeLib dll
                            CropImageByRect(pngFile, regionOfBarcode, barcodePath);

                            var barcodeLibRightPartStrings = new List<string>();
                            if (result.CodeType.ToString() == "QR")
                            {

                                var QRRes = QRCodeDecoderLib(barcodePath);
                                if (QRRes.Count > 1)
                                {
                                    ;
                                }
                                resultValue = QRRes.FirstOrDefault();
                            }
                            else
                            {
                                //barcodeLibRightPartStrings.AddRange(BarcodeLibBarcodeReader(barcodePath));
                                barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));
                                resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);
                            }




                            //Зачищаем после обработки временный файл
                            File.Delete(barcodePath);

                        }
                        else
                        {
                            //Сразу забираем значение для типов CODE39 и CODE39EX, т.к. бесплатная версия ASPOSE их не режет
                            resultValue = asposeDirtyRes;

                        }

                        BarcodeNumPage = pdfPageNum;
                        BarcodeValuesAndTypes.Add(new Tuple<string, string>(result.CodeType.ToString(), resultValue));


                        //KeepAutomationBarcodeReader(barcodePath);



                    }
                    bitmap.Dispose();
                }





                //выходим если нашли номера на текущей странице
                if (BarcodeValuesAndTypes.Count > 0) break;
            }
        }

        /// <summary>
        /// Получает значения баркодов с первой успешно распознанной страницы
        /// </summary>
        /// <param name="pngfiles"></param>
        /// <param name="heightPersentageArea">Область поиска по высоте листа в процентах</param>
        private void GetBarcodeValuesAndType4libs(List<string> pngfiles, int heightPersentageArea)
        {
            foreach (var pngFile in pngfiles)
            {
               
                var bitmap = new Bitmap(pngFile);

                using (Aspose.BarCode.BarCodeRecognition.BarCodeReader reader = new Aspose.BarCode.BarCodeRecognition.BarCodeReader(bitmap))
                {

                    //если страница достаточно большая то для ускорения задаем поисковую область
                    if (bitmap.Height > 500)
                    {

                        var findArea = new Rectangle(0, 0, bitmap.Width, bitmap.Height / 100 * heightPersentageArea);
                        //var outputFile = Path.Combine(Path.GetDirectoryName(pngFile), String.Concat(Path.GetRandomFileName(), @".png"));
                        //CropImageByRect(pngFile, findArea, outputFile);
                        reader.SetBarCodeImage(bitmap, findArea);
                    }


                    var readerRes = reader.ReadBarCodes();


                    foreach (Aspose.BarCode.BarCodeRecognition.BarCodeResult result in readerRes)
                    {

                        int pdfPageNum = GetNumPageFromPath(pngFile);
                        //Console.WriteLine("Symbology Type: " + result.CodeType);
                        //Console.WriteLine("CodeText: " + result.CodeText);

                        var asposeDirtyRes = result.CodeText;
                        var matchWatermarkRes = Regex.Match(asposeDirtyRes, @"(.{1,})(Recognized)");
                        var resultValue = "";
                        //Если нашли Watermark, то будем получать результат совместно из 3 библиотек
                        if (matchWatermarkRes.Success)
                        {
                            var asposeBarcodePart = matchWatermarkRes.Groups[1].Value;

                            //область где найден баркод
                            var findedBarcodeRegion = result.Region.Rectangle;
                            var barcodeType = result.CodeType.ToString();

                            resultValue = ExtendedBarcodeAnalize(pngFile, asposeBarcodePart, barcodeType, findedBarcodeRegion);


                            #region oldCodeCommented
                            //var regionOfBarcode = new Rectangle(findedBarcodeRegion.X - 10, findedBarcodeRegion.Y - 10, findedBarcodeRegion.Width + 20, findedBarcodeRegion.Height + 20);
                            //var barcodePath = Path.Combine(Path.GetDirectoryName(pngFile), String.Concat(Path.GetRandomFileName(), @".png"));

                            ////Сохраняем только сам баркод под распознование SpireBarcode
                            //CropImageByRect(pngFile, regionOfBarcode, barcodePath);


                            //resultValue = FreeSpireLib(barcodePath).FirstOrDefault();

                            ////resultValue = FreeSpireLib(pngFile).FirstOrDefault();


                            //if (resultValue == "0" || resultValue == "" || resultValue == null)
                            //{
                            //    var barcodeLibRightPartStrings = new List<string>();
                            //    if (result.CodeType.ToString() == "QR")
                            //    {

                            //        var QRRes = QRCODEDECODERLIB(barcodePath);
                            //        //var QRRes1 = QRCODEDECODERLIB(pngFile);
                            //        //if (QRRes.Count > 1)
                            //        //{
                            //        //    ;
                            //        //}
                            //        resultValue = QRRes.FirstOrDefault();
                            //        if (resultValue == null || resultValue == "")
                            //        {
                            //            barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));

                            //            resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);

                            //            if (resultValue == null || resultValue == "")
                            //            {
                            //                resultValue = asposeBarcodePart;
                            //            }
                            //        }

                            //    }
                            //    else
                            //    {
                            //        //barcodeLibRightPartStrings.AddRange(BarcodeLibBarcodeReader(barcodePath));
                            //        //barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(pngFile));
                            //        barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));

                            //        resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);
                            //    }
                            //}


                            ////Зачищаем после обработки временный файл
                            //File.Delete(barcodePath);
                            #endregion

                        }
                        else
                        {
                            //Сразу забираем значение для типов CODE39 и CODE39EX, т.к. бесплатная версия ASPOSE их не режет
                            resultValue = asposeDirtyRes;

                        }

                        BarcodeNumPage = pdfPageNum;
                        BarcodeValuesAndTypes.Add(new Tuple<string, string>(result.CodeType.ToString(), resultValue));


                        //KeepAutomationBarcodeReader(barcodePath);



                    }

                    //диспозим битмап,т.к. в многопотоке он не успевает освободиться и выдает ошибку при попытке удаления файла
                    bitmap.Dispose();
                }





                //выходим если нашли номера на текущей странице
                if (BarcodeValuesAndTypes.Count > 0) break;
            }
        }

        private void GetBarcodeValuesAndType4libs(Dictionary<int, MemoryStream> pngfiles, string outputDir, int heightPersentageArea)
        {
            foreach (var pngFile in pngfiles)
            {
                //Stream stream = new MemoryStream(File.ReadAllBytes(pngFile));
                var bitmap = new Bitmap(pngFile.Value);

                using (Aspose.BarCode.BarCodeRecognition.BarCodeReader reader = new Aspose.BarCode.BarCodeRecognition.BarCodeReader(bitmap))
                {

                    //если страница достаточно большая то для ускорения задаем поисковую область
                    if (bitmap.Height > 500)
                    {

                        var findArea = new Rectangle(0, 0, bitmap.Width, bitmap.Height / 100 * heightPersentageArea);
                        //var outputFile = Path.Combine(Path.GetDirectoryName(pngFile), String.Concat(Path.GetRandomFileName(), @".png"));
                        //CropImageByRect(pngFile, findArea, outputFile);
                        reader.SetBarCodeImage(bitmap, findArea);
                    }


                    var readerRes = reader.ReadBarCodes();


                    foreach (Aspose.BarCode.BarCodeRecognition.BarCodeResult result in readerRes)
                    {

                        int pdfPageNum = pngFile.Key;
                        //Console.WriteLine("Symbology Type: " + result.CodeType);
                        //Console.WriteLine("CodeText: " + result.CodeText);

                        var asposeDirtyRes = result.CodeText;
                        var matchWatermarkRes = Regex.Match(asposeDirtyRes, @"(.{1,})(Recognized)");
                        var resultValue = "";
                        //Если нашли Watermark, то будем получать результат совместно из 3 библиотек
                        if (matchWatermarkRes.Success)
                        {
                            var asposeBarcodePart = matchWatermarkRes.Groups[1].Value;

                            //область где найден баркод
                            var findedBarcodeRegion = result.Region.Rectangle;
                            var barcodeType = result.CodeType.ToString();

                            resultValue = ExtendedBarcodeAnalize(pngFile.Value, outputDir, asposeBarcodePart, barcodeType, findedBarcodeRegion);




                        }
                        else
                        {
                            //Сразу забираем значение для типов CODE39 и CODE39EX, т.к. бесплатная версия ASPOSE их не режет
                            resultValue = asposeDirtyRes;

                        }

                        BarcodeNumPage = pdfPageNum;
                        BarcodeValuesAndTypes.Add(new Tuple<string, string>(result.CodeType.ToString(), resultValue));


                        //KeepAutomationBarcodeReader(barcodePath);



                    }

                    //диспозим битмап,т.к. в многопотоке он не успевает освободиться и выдает ошибку при попытке удаления файла
                    bitmap.Dispose();
                }





                //выходим если нашли номера на текущей странице
                if (BarcodeValuesAndTypes.Count > 0) break;
            }
        }




        private static string ExtendedBarcodeAnalize(string orignPng, string asposeBarcodePart, string barcodeType, Rectangle findedBarcodeRegion)
        {
            string resultValue = "";

            var regionOfBarcode = new Rectangle(findedBarcodeRegion.X - 10, findedBarcodeRegion.Y - 10, findedBarcodeRegion.Width + 20, findedBarcodeRegion.Height + 20);
            var barcodePath = Path.Combine(Path.GetDirectoryName(orignPng), String.Concat(Path.GetRandomFileName(), @".png"));

            //Сохраняем только сам баркод под распознование SpireBarcode
            CropImageByRect(orignPng, regionOfBarcode, barcodePath);


            resultValue = FreeSpireLib(barcodePath).FirstOrDefault();




            if (resultValue == "0" || resultValue == "" || resultValue == null)
            {
                var barcodeLibRightPartStrings = new List<string>();
                if (barcodeType == "QR")
                {

                    var QRRes = QRCodeDecoderLib(barcodePath);

                    resultValue = QRRes.FirstOrDefault();
                    if (resultValue == null || resultValue == "")
                    {
                        barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));

                        resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);

                        if (resultValue == null || resultValue == "")
                        {
                            resultValue = asposeBarcodePart;
                        }
                    }

                }
                else
                {
                    //barcodeLibRightPartStrings.AddRange(BarcodeLibBarcodeReader(barcodePath));
                    //barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(pngFile));
                    barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));

                    resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);
                }
            }


            //Зачищаем после обработки временный файл
            File.Delete(barcodePath);
            return resultValue;

        }


        private static string ExtendedBarcodeAnalize(MemoryStream orignPng, string outPutDir, string asposeBarcodePart, string barcodeType, Rectangle findedBarcodeRegion)
        {
            string resultValue = "";

            var regionOfBarcode = new Rectangle(findedBarcodeRegion.X - 10, findedBarcodeRegion.Y - 10, findedBarcodeRegion.Width + 20, findedBarcodeRegion.Height + 20);
            var barcodePath = Path.Combine(outPutDir, String.Concat(Path.GetRandomFileName(), @".png"));

            //Сохраняем только сам баркод под распознование SpireBarcode
            CropImageByRect(orignPng, regionOfBarcode, barcodePath);


            resultValue = FreeSpireLib(barcodePath).FirstOrDefault();


            if (resultValue == "0" || resultValue == "" || resultValue == null)
            {
                var barcodeLibRightPartStrings = new List<string>();
                if (barcodeType == "QR")
                {

                    var QRRes = QRCodeDecoderLib(barcodePath);

                    resultValue = QRRes.FirstOrDefault();
                    if (resultValue == null || resultValue == "")
                    {
                        barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));
                        //barcodeLibRightPartStrings.AddRange(BarcodeLibBarcodeReader(barcodePath));
                        resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);

                        if (resultValue == null || resultValue == "")
                        {
                            resultValue = asposeBarcodePart;
                        }
                    }

                }
                else
                {
                    //barcodeLibRightPartStrings.AddRange(BarcodeLibBarcodeReader(barcodePath));
                    
                    barcodeLibRightPartStrings.AddRange(KeepAutomationBarcodeReader(barcodePath));

                    resultValue = GetFullBarcodeValue(asposeBarcodePart, barcodeLibRightPartStrings);
                }
            }


            //Зачищаем после обработки временный файл
            File.Delete(barcodePath);
            return resultValue;

        }


        /// <summary>
        /// Совмещаем результат значений баркодов из двух библиотек, в итоге получаем корректное значение
        /// </summary>
        /// <param name="leftPart"></param>
        /// <param name="rightPart"></param>
        /// <returns></returns>
        private static string GetFullBarcodeValue(string leftPart, IEnumerable<string> rightPart)
        {
            var result = "";
            var leftPartWithoutStars = leftPart.Replace("*", "");


            //var temp = rightPart.FirstOrDefault().Substring(1);
            //var temp2= leftPartWithoutStars.Substring(1);

            //var rightPartRes = rightPart.Where(str => leftPart.Contains(str.Substring(1, str.Length - (1+ 2)))).FirstOrDefault();
            var rightPartRes = rightPart.Where(str => str.Substring(1).Contains(leftPartWithoutStars.Substring(1))).FirstOrDefault();
            if (rightPartRes != null)
            {
                result = String.Concat(leftPart.Substring(0, 1), rightPartRes.Substring(1));
            }
            else
            {
                result = String.Format("Ошибка слияния значений : {0} и {1}", leftPart, String.Join(";", rightPart));
            }
            return result;
        }

        /// <summary>
        /// Вырезает изображение из указанной области
        /// </summary>
        /// <param name="inPath">Оригинал изображения</param>
        /// <param name="cropRect">Область выделения</param>
        /// <param name="outPath">Новый файл png</param>
        public static void CropImageByRect(string inPath, Rectangle cropRect, string outPath)
        {
            byte[] photoBytes = File.ReadAllBytes(inPath);

            ImageProcessor.Imaging.Formats.ISupportedImageFormat format = new ImageProcessor.Imaging.Formats.PngFormat { Quality = 100 };

            using (MemoryStream inStream = new MemoryStream(photoBytes))
            {

                // Initialize the ImageFactory using the overload to preserve EXIF metadata.
                using (ImageProcessor.ImageFactory imageFactory = new ImageProcessor.ImageFactory(preserveExifData: true))
                {
                    // Load, resize, set the format and quality and save an image.
                    imageFactory.Load(inStream)
                                .Crop(cropRect)
                                .Format(format)
                                .Save(outPath);

                }

            }

        }

        public static void CropImageByRect(MemoryStream inPath, Rectangle cropRect, string outPath)
        {
            byte[] photoBytes = inPath.ToArray();

            ImageProcessor.Imaging.Formats.ISupportedImageFormat format = new ImageProcessor.Imaging.Formats.PngFormat { Quality = 100 };

            using (MemoryStream inStream = new MemoryStream(photoBytes))
            {

                // Initialize the ImageFactory using the overload to preserve EXIF metadata.
                using (ImageProcessor.ImageFactory imageFactory = new ImageProcessor.ImageFactory(preserveExifData: true))
                {
                    // Load, resize, set the format and quality and save an image.
                    imageFactory.Load(inStream)
                                .Crop(cropRect)
                                .Format(format)
                                .Save(outPath);

                }

            }

        }



        private static List<string> FreeSpireLib(string page1FileName)
        {
            var datas = new List<string>();

            Spire.Barcode.Settings.BarcodeInfo[] res = new Spire.Barcode.Settings.BarcodeInfo[1];

            bool spireIsCorrect = true;

            while (spireIsCorrect)
            {
                Task<Spire.Barcode.Settings.BarcodeInfo[]> taskBarcodeRead = new Task<Spire.Barcode.Settings.BarcodeInfo[]>(() => Spire.Barcode.BarcodeScanner.ScanInfo(page1FileName));
                taskBarcodeRead.Start();
                taskBarcodeRead.Wait();
                res = taskBarcodeRead.Result;
                if (res.Count() > 0)
                {
                    foreach (var r in res)
                    {
                        if (r != null)
                        {
                            
                            datas.Add(r.DataString);
                            spireIsCorrect = false;
                        }
              
                    }
                }
                else
                {
                    spireIsCorrect = false;
                }


            }


            return datas;
        }


        private static String[] BarcodeLibBarcodeReader(string path)
        {
            BarcodeLib.BarcodeReader.OptimizeSetting setting = new BarcodeLib.BarcodeReader.OptimizeSetting();

            string[] datas = new string[] {""};

            setting.setMaxOneBarcodePerPage(true);
          
            try
            {
                //var reader = BarcodeLib.BarcodeReader.BarcodeReader.ReadBarcode(path, new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 });
                //datas = reader.Select(d => d.Data).ToArray();
                datas = BarcodeLib.BarcodeReader.BarcodeReader.read(path, new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 }, setting);
            }
            catch
            {
                
            }

            //String[] datas = BarcodeLib.BarcodeReader.BarcodeReader.read(path, new int[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 });

            //Console.WriteLine("BarcodeLibBarcodeReader: ");
            //foreach (var res in datas)
            //{
            //    Console.WriteLine(res);
            //}

            return datas;
        }


        public static List<string> KeepAutomationBarcodeReader(string path)
        {
            List<string> datas = new List<string>();
            //try
            //{
                var res = KeepAutomation.BarcodeReader.BarcodeReader.readBarcode(path, new KeepAutomation.BarcodeReader.BarcodeType[]
                { KeepAutomation.BarcodeReader.BarcodeType.Code128, KeepAutomation.BarcodeReader.BarcodeType.Codabar,
                  KeepAutomation.BarcodeReader.BarcodeType.Code39, KeepAutomation.BarcodeReader.BarcodeType.Code39Ex,
                  KeepAutomation.BarcodeReader.BarcodeType.EAN8,
                  KeepAutomation.BarcodeReader.BarcodeType.EAN13,
                  KeepAutomation.BarcodeReader.BarcodeType.Interleaved25 ,
                  KeepAutomation.BarcodeReader.BarcodeType.UPCA ,
                  KeepAutomation.BarcodeReader.BarcodeType.UPCE,
                  KeepAutomation.BarcodeReader.BarcodeType.DataMatrix,
                  KeepAutomation.BarcodeReader.BarcodeType.PDF417 ,
                  KeepAutomation.BarcodeReader.BarcodeType.QRCode
                });


                if (res != null)
                {
                    foreach (var r in res)
                    {
                        datas.Add(r.CodeToEncode);
                    }
                }
            //}
            //catch
            //{

            //}

            return datas;

        }


        public static List<string> QRCodeDecoderLib(string path)
        {

            List<string> datas = new List<string>();
            QRCodeDecoderLibrary.QRDecoder decoder = new QRCodeDecoderLibrary.QRDecoder();

            Bitmap bitmap = new Bitmap(path);

            var resultArray = decoder.ImageDecoder(bitmap);

            if (resultArray != null)
            {
                foreach (var t in resultArray)
                {
                    var str = ByteArrayToStr(t);
                    datas.Add(str);
                }
            }


            bitmap.Dispose();
            return datas;
        }


        public static string ByteArrayToStr(byte[] DataArray)
        {
            var Decoder = Encoding.UTF8.GetDecoder();
            int CharCount = Decoder.GetCharCount(DataArray, 0, DataArray.Length);
            char[] CharArray = new char[CharCount];
            Decoder.GetChars(DataArray, 0, DataArray.Length, CharArray, 0);
            return new string(CharArray);
        }


        private static int GetNumPageFromPath(string fileNameWithPage)
        {
            var fileName = Path.GetFileName(fileNameWithPage);
            var pageNum = Regex.Match(fileName, @"_page([0-9]{1,})\.").Groups[1].Value;
            return Int32.Parse(pageNum);
        }

        public static List<string> SplitBySinglePagesMethod(string outputFolder, string baseFilename, string pageFilenameAnnex = "_page")
        {
            List<string> pagesFilenames = new List<string>();
            var baseFileNameWithoutExst = Path.GetFileNameWithoutExtension(baseFilename);

            var Filestream = new MemoryStream(File.ReadAllBytes(baseFilename))
            {
                Position = 0
            };

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    var _lock = new object();
                    lock (_lock)
                    {
                        Filestream.CopyTo(ms);
                        ms.Position = 0;
                    }

                    var currentConsoleColor = Console.ForegroundColor;
                    //Console.ForegroundColor = ConsoleColor.Gray;
                    //Console.WriteLine($"Splitting file => '{baseFilename}'");
                    using (iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(ms)))
                    {
                        int totalPages = pdfDoc.GetNumberOfPages();

                        for (int i = 1; i <= totalPages; i++)
                        {
                            try
                            {
                                string path = outputFolder + $"\\{baseFileNameWithoutExst}{pageFilenameAnnex}" + i.ToString() + ".pdf";
                                pagesFilenames.Add(path);

                                using (iText.Kernel.Pdf.PdfWriter wr = new iText.Kernel.Pdf.PdfWriter(path))
                                {
                                    using (iText.Kernel.Pdf.PdfDocument pdfDoc1 = new iText.Kernel.Pdf.PdfDocument(wr))
                                    {
                                        pdfDoc.CopyPagesTo(i, i, pdfDoc1);
                                        pdfDoc1.Close();
                                    }
                                }
                            }
                            catch (NullReferenceException ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                            catch (ArgumentNullException ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                            catch (SystemException ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                        }
                    }
                }
            }
            catch (NullReferenceException ex)
            {
                throw new Exception("Pdf split error", ex);
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception("Pdf split error", ex);
            }
            catch (SystemException ex)
            {
                throw new Exception("Pdf split error", ex);
            }
            catch
            {
                throw new Exception("Pdf split error");
            }

            return pagesFilenames;
        }

        public static Dictionary<int, byte[]> SplitBySinglePagesMethod(string outputFolder, string baseFilename, MemoryStream baseFilenameStream, string pageFilenameAnnex = "_page")
        {
            Dictionary<int, byte[]> pagesFileStreams = new Dictionary<int, byte[]>();
            var baseFileNameWithoutExst = Path.GetFileNameWithoutExtension(baseFilename);



            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    var _lock = new object();
                    lock (_lock)
                    {
                        baseFilenameStream.CopyTo(ms);
                        ms.Position = 0;
                    }

                    var currentConsoleColor = Console.ForegroundColor;

                    using (iText.Kernel.Pdf.PdfDocument pdfDoc = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(ms)))
                    {
                        int totalPages = pdfDoc.GetNumberOfPages();

                        for (int i = 1; i <= totalPages; i++)
                        {
                            try
                            {



                                using (MemoryStream memStream = new MemoryStream())
                                {
                                    using (iText.Kernel.Pdf.PdfWriter pdfWriter = new iText.Kernel.Pdf.PdfWriter(memStream))
                                    {
                                        pdfWriter.SetCloseStream(true);
                                        using (iText.Kernel.Pdf.PdfDocument pdfDoc1 = new iText.Kernel.Pdf.PdfDocument(pdfWriter))
                                        {
                                            pdfDoc.CopyPagesTo(i, i, pdfDoc1);
                                            pdfDoc1.Close();
                                        }
                                    }
                                    var bytes = memStream.ToArray();
                                    pagesFileStreams.Add(i, bytes);
                                }





                            }
                            catch (NullReferenceException ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                            catch (ArgumentNullException ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                            catch (SystemException ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Pdf split error", ex);
                            }
                        }
                    }
                }
            }
            catch (NullReferenceException ex)
            {
                throw new Exception("Pdf split error", ex);
            }
            catch (ArgumentNullException ex)
            {
                throw new Exception("Pdf split error", ex);
            }
            catch (SystemException ex)
            {
                throw new Exception("Pdf split error", ex);
            }
            catch
            {
                throw new Exception("Pdf split error");
            }

            return pagesFileStreams;
        }

    }
}
