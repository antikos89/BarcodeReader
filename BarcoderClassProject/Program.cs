using BarcoderClassProject.Helpers;
using NDesk.Options;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BarcoderClassProject
{
    internal class Program
    {
        static void Main(string[] args)
        {

            //Путь до папки с файлами документации
            string folderWithPdf = "";
            //временная директория
            string tempFolderPath = "";
            //Имя ексель файла для сохранения полученных данныых
            string excellOutputFileNameWithoutExt = "";
            //количество потоков
            int maxThreads = 4;
            //Разрешение при переводе PDF в PNG, работает по типу настроек сканера 300 dpi дает самые хорошие результаты
            int scanResolution = 300;

            //Область в которой ищем баркод на листе вычислиятся по высоте указанных в процентах
            int heightPersentageArea = 10;
            //Использовать память для хранения всех файлов или сохранять файлы в файловую систему
            var useMemoryToloadAllFiles = true;


            var startArgs = GetStartParams(args);
            folderWithPdf = startArgs.Item1;
            tempFolderPath = startArgs.Item2;
            excellOutputFileNameWithoutExt = startArgs.Item3;
            //maxThreads = Environment.ProcessorCount / 2;

            maxThreads = startArgs.Item4;
            scanResolution = startArgs.Item5;
            heightPersentageArea = startArgs.Item6;
            useMemoryToloadAllFiles = startArgs.Item7;




            if (useMemoryToloadAllFiles)
            {
                //ProcessPDFFolderWithParallelTest(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderWithMemoryStream(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderWithParallel(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderWithTask(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderWithSemaphore(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                ProcessPDFFolderWithSemaphoreAndStreams(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderWithTaskAlt(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderWithTaskAltWithStreams(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
                //ProcessPDFFolderParallelFor(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
            }
            else
            {
                ProcessPDFFolderWithSemaphore(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreads, scanResolution, heightPersentageArea);
            }

            ConsoleAwesome.WriteLine("Обработка всех файлов завершена", ConsoleColor.White);
            Console.ReadLine();
        }


        public static void TestBarcodereader()
        {
            var path = @"C:\Users\kosPC\Desktop\работа\Проекты\AMURGCCPROJECT(nextTask)\TestsFolder\testFile\TMP-002-P14.pdf";
            var barcodereader = new CustomBarcodeReader(path);
        }


        public static void ProcessPDFFolderWithParallelTest(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();
            var pdfFilesList2 = Directory.GetFiles(@"C:\Users\kosPC\Desktop\работа\Проекты\Фазлеев(Barcodes)\TestsFolder\test200", @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();

            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();



            //;
            ThreadPool.SetMinThreads(4000, 100);
            ThreadPool.SetMaxThreads(4000, 100);

            //Parallel.Invoke(
            //        () =>
            //        {
            //            ThreadPoolLibrary.FixedThreadPool pool = new ThreadPoolLibrary.FixedThreadPool(16);

            //            bool added = pool.Execute(new ThreadPoolLibrary.Task(() =>
            //            {
            //                var newList = new ConcurrentQueue<string>();
            //                pdfFilesList/*.Take(50).ToList()*/.ForEach(j => newList.Enqueue(j));

            //                Parallel.ForEach<string>(newList, new ParallelOptions { MaxDegreeOfParallelism = 16 }, pdfFile =>
            //                {
            //                    var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                    var u = result.Reverse();
            //                    u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //                });
            //                //newList.AsParallel().WithDegreeOfParallelism(10).ForAll(pdfFile =>
            //                //{


            //                //    //////////////////////////////////////////////////////////
            //                //    ////общая обработка
            //                //    //////////////////////////////////////////////////////////
            //                //    //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                //    ////result.Reverse();
            //                //    //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //                //    var result = ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                //    var u = result.Result.Reverse();
            //                //    u.ToList().ForEach(j => globalBarcodeItems.Add(j));


            //                //});
            //            }));
            //            pool.Stop();


            //        },
            //        () =>
            //        {
            //            ThreadPoolLibrary.FixedThreadPool pool1 = new ThreadPoolLibrary.FixedThreadPool(16);

            //            bool added = pool1.Execute(new ThreadPoolLibrary.Task(() =>
            //            {
            //                var newList2 = new ConcurrentQueue<string>();
            //                pdfFilesList2./*Skip(50).Take(50).ToList().*/ForEach(i => newList2.Enqueue(i));


            //                Parallel.ForEach<string>(newList2, new ParallelOptions { MaxDegreeOfParallelism = 16 }, pdfFile =>
            //                {
            //                    var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                    var u = result.Reverse();
            //                    u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //                });

            //    newList2.AsParallel().WithDegreeOfParallelism(10).ForAll(pdfFile =>
            //    {




            //        //////////////////////////////////////////////////////////
            //        ////общая обработка
            //        //////////////////////////////////////////////////////////
            //        //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //        ////result.Reverse();
            //        //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //        var result = ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //        var u = result.Result.Reverse();
            //        u.ToList().ForEach(j => globalBarcodeItems.Add(j));


            //    //    });
            //}));
            //                pool1.Stop();
            //            }
            //        );




            //var t = 0;
            //var t2 = 0;
            //ThreadPool.GetAvailableThreads(out t,out t2);
            //;
            //ThreadPool.SetMinThreads(100, 100);


            //var semaphoreSlim = new SemaphoreSlim(16);
            //var semaphoreSlim2 = new SemaphoreSlim(16);
            //var newList = new List<string>();
            //newList.AddRange(pdfFilesList.GetRange(0, pdfFilesList.Count / 2));

            //var newList2 = new List<string>();
            //newList2.AddRange(pdfFilesList.GetRange(pdfFilesList.Count / 2, pdfFilesList.Count / 2));
            //var tasks = new List<Task>(newList.Count);
            //var tasks2 = new List<Task>(newList2.Count);

            //foreach (var pdfFile in newList)
            //{
            //    tasks.Add(Task.Run(() =>
            //    {
            //        semaphoreSlim.Wait();
            //        try
            //        {
            //            //////////////////////////////////////////////////////////
            //            ////общая обработка
            //            var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);

            //            //result.Reverse();
            //            result.ToList().ForEach(i => globalBarcodeItems.Add(i));



            //        }
            //        finally
            //        {
            //            semaphoreSlim.Release();
            //        }
            //    }));
            //}

            //foreach (var pdfFile in newList2)
            //{
            //    tasks2.Add(Task.Run(() =>
            //    {
            //        semaphoreSlim2.Wait();
            //        try
            //        {
            //            //////////////////////////////////////////////////////////
            //            ////общая обработка
            //            var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);

            //            //result.Reverse();
            //            result.ToList().ForEach(i => globalBarcodeItems.Add(i));



            //        }
            //        finally
            //        {
            //            semaphoreSlim2.Release();
            //        }
            //    }));
            //}

            //Task.WaitAll(tasks.ToArray());
            //Task.WaitAll(tasks2.ToArray());



            //Parallel.For(0, pdfFilesList.Count, new ParallelOptions { MaxDegreeOfParallelism = -1 }, i =>
            //{
            //    //////////////////////////////////////////////////////////
            //    ////общая обработка
            //    //////////////////////////////////////////////////////////

            //    var result = ReaderResultItem.ProcessingPdfFile(pdfFilesList[i], tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    var u = result.Reverse();
            //    u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //});

            //var task1 = Task.Run(() =>
            //    {
            //        var newList = new List<string>();
            //        newList.AddRange(pdfFilesList.GetRange(0, pdfFilesList.Count / 2));
            //        Parallel.For(0, newList.Count, new ParallelOptions { MaxDegreeOfParallelism = 20 }, i =>
            //        {
            //        //////////////////////////////////////////////////////////
            //        ////общая обработка
            //        //////////////////////////////////////////////////////////

            //            var result = ReaderResultItem.ProcessingPdfFileTask(newList[i], tempDir, maxThreads, scanResolution, heightPersentageArea);
            //            var u = result.Result.Reverse();
            //            u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //        });
            //    }
            //);

            //var task2 = Task.Run(() =>
            //{
            //    var newList2 = new List<string>();
            //    newList2.AddRange(pdfFilesList.GetRange(pdfFilesList.Count / 2, pdfFilesList.Count / 2));
            //    Parallel.For(0, newList2.Count, new ParallelOptions { MaxDegreeOfParallelism = 20 }, i =>
            //    {
            //        //////////////////////////////////////////////////////////
            //        ////общая обработка
            //        //////////////////////////////////////////////////////////

            //        var result = ReaderResultItem.ProcessingPdfFileTask(newList2[i], tempDir, maxThreads, scanResolution, heightPersentageArea);
            //        var u = result.Result.Reverse();
            //        u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //    });
            //}
            //);

            //Task.WaitAll(task1, task2);

            //var halfList = pdfFilesList.Count / 2;
            //Parallel.Invoke(
            //        () =>
            //        {
            //            //var newList = new List<string>();
            //            //newList.AddRange(pdfFilesList.GetRange(0, pdfFilesList.Count / 2));
            //            var newList = new ConcurrentQueue<string>();
            //            pdfFilesList.Take(50).ToList().ForEach(i => newList.Enqueue(i));

            //            newList.AsParallel().WithDegreeOfParallelism(12).ForAll(pdfFile =>
            //            {


            //                //////////////////////////////////////////////////////////
            //                ////общая обработка
            //                //////////////////////////////////////////////////////////
            //                //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                ////result.Reverse();
            //                //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //                var result = ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                var u = result.Result.Reverse();
            //                u.ToList().ForEach(i => globalBarcodeItems.Add(i));


            //            });

            //            //Parallel.For(0, newList.Count, new ParallelOptions { MaxDegreeOfParallelism = -1 }, i =>
            //            //{
            //            //    //////////////////////////////////////////////////////////
            //            //    ////общая обработка
            //            //    //////////////////////////////////////////////////////////

            //            //    var result = ReaderResultItem.ProcessingPdfFileTask(newList[i], tempDir, maxThreads, scanResolution, heightPersentageArea);
            //            //    var u = result.Result.Reverse();
            //            //    u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //            //});


            //        },
            //        () =>
            //        {
            //            //var newList2 = new List<string>();
            //            //newList2.AddRange(pdfFilesList.GetRange(pdfFilesList.Count / 2, pdfFilesList.Count / 2));
            //            var newList2 = new ConcurrentQueue<string>();
            //            pdfFilesList.Skip(50).Take(50).ToList().ForEach(i => newList2.Enqueue(i));

            //            newList2.AsParallel().WithDegreeOfParallelism(12).ForAll(pdfFile =>
            //            {


            //                //////////////////////////////////////////////////////////
            //                ////общая обработка
            //                //////////////////////////////////////////////////////////
            //                //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                ////result.Reverse();
            //                //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //                var result = ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                var u = result.Result.Reverse();
            //                u.ToList().ForEach(i => globalBarcodeItems.Add(i));


            //            });
            //            //Parallel.For(0, newList2.Count, new ParallelOptions { MaxDegreeOfParallelism = -1 }, i =>
            //            //{
            //            //    //////////////////////////////////////////////////////////
            //            //    ////общая обработка
            //            //    //////////////////////////////////////////////////////////

            //            //    var result = ReaderResultItem.ProcessingPdfFileTask(newList2[i], tempDir, maxThreads, scanResolution, heightPersentageArea);
            //            //    var u = result.Result.Reverse();
            //            //    u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //            //});
            //        }
            //    );



            //pdfFilesList.AsParallel().WithDegreeOfParallelism(12).WithMergeOptions(ParallelMergeOptions.FullyBuffered).ForAll(pdfFile =>
            //{


            //    //////////////////////////////////////////////////////////
            //    ////общая обработка
            //    //////////////////////////////////////////////////////////
            //    //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    //result.Reverse();
            //    //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //    var result = ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    var u = result.Result.Reverse();
            //    u.ToList().ForEach(i => globalBarcodeItems.Add(i));


            //});




            //Parallel.Invoke(
            //        () =>
            //        {

            //                var newList = new ConcurrentQueue<string>(pdfFilesList);


            //                Parallel.ForEach<string>(newList, new ParallelOptions { MaxDegreeOfParallelism = 16 }, pdfFile =>
            //                {
            //                    Task.Run(() =>
            //                    {
            //                        var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                        var u = result.Reverse();
            //                        u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //                    });

            //                });



            //        },
            //        () =>
            //        {

            //                var newList2 = new ConcurrentQueue<string>(pdfFilesList2);



            //                Parallel.ForEach<string>(newList2, new ParallelOptions { MaxDegreeOfParallelism = 16 }, pdfFile =>
            //                {
            //                    Task.Run(() =>
            //                    {
            //                        var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //                        var u = result.Reverse();
            //                        u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            //                    });
            //                });


            //        }
            //    );




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }



        public static void ProcessPDFFolderWithParallel(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///


            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();

            Stopwatch timerMemoryStream = new Stopwatch();
            timerMemoryStream.Start();
            double totalTimeM = 0;
            double totalSecondsM = 0.0f;
            string totalSecondsValueM = "";

            ConcurrentDictionary<string, MemoryStream> memoryStreams = new ConcurrentDictionary<string, MemoryStream>();
            pdfFilesList.AsParallel().ForAll(pdfFile =>
            {
                var memoryStream = new MemoryStream(File.ReadAllBytes(pdfFile))
                {
                    Position = 0
                };

                memoryStreams.TryAdd(pdfFile, memoryStream);
            });



            timerMemoryStream.Stop();
            totalSecondsM = timerMemoryStream.Elapsed.TotalSeconds;
            totalTimeM += totalSeconds;
            totalSecondsValueM = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На загрузку всех файлов в память затраченно = " + totalSecondsM + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            //pdfFilesList.AsParallel().WithDegreeOfParallelism(24).WithMergeOptions(ParallelMergeOptions.FullyBuffered).ForAll(pdfFile =>
            //{


            //    //////////////////////////////////////////////////////////
            //    ////общая обработка
            //    //////////////////////////////////////////////////////////
            //    //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    //result.Reverse();
            //    //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //    var result = ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    var u = result.Result.Reverse();
            //    u.ToList().ForEach(i => globalBarcodeItems.Add(i));


            //});

            //ThreadPool.SetMinThreads(12, 100);
            //ThreadPool.SetMaxThreads(24, 100);

            var arr = memoryStreams.ToArray();

            Parallel.ForEach<KeyValuePair<string, MemoryStream>>(arr, new ParallelOptions { MaxDegreeOfParallelism = 20 }, oneFile =>
            {

                var pdfFileName = oneFile.Key;
                var pdfFileStream = oneFile.Value;

                var result = ReaderResultItem.ProcessingPdfFileTask(pdfFileName, pdfFileStream, tempDir, maxThreads, scanResolution, heightPersentageArea, false);
                result.Result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            });


            //memoryStreams.AsParallel().WithDegreeOfParallelism(24).WithMergeOptions(ParallelMergeOptions.FullyBuffered).ForAll(memoryStream => {
            //    var pdfFileName = memoryStream.Key;
            //    var pdfFileStream= memoryStream.Value;

            //    var result = ReaderResultItem.ProcessingPdfFileTask(pdfFileName, pdfFileStream, tempDir, maxThreads, scanResolution, heightPersentageArea,false);
            //    result.Result.ToList().ForEach(i => globalBarcodeItems.Add(i));


            //});


            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }



        public static void ProcessPDFFolderWithMemoryStream(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();


            Stopwatch timerMemoryStream = new Stopwatch();
            timerMemoryStream.Start();
            double totalTimeM = 0;
            double totalSecondsM = 0.0f;
            string totalSecondsValueM = "";

            ConcurrentDictionary<string, MemoryStream> memoryStreams = new ConcurrentDictionary<string, MemoryStream>();
            pdfFilesList.AsParallel().ForAll(pdfFile =>
            {
                var memoryStream = new MemoryStream(File.ReadAllBytes(pdfFile))
                {
                    Position = 0
                };

                memoryStreams.TryAdd(pdfFile, memoryStream);
            });



            timerMemoryStream.Stop();
            totalSecondsM = timerMemoryStream.Elapsed.TotalSeconds;
            totalTimeM += totalSeconds;
            totalSecondsValueM = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На загрузку всех файлов в память затраченно = " + totalSecondsM + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);




            memoryStreams.AsParallel().ForAll(pdfFile =>
            {
                var result = ReaderResultItem.ProcessingPdfFile(pdfFile.Key, pdfFile.Value, tempDir, maxThreads, scanResolution, heightPersentageArea);
                result.ToList().ForEach(j => globalBarcodeItems.Add(j));
            });




            //List<Task<IEnumerable<ReaderResultItem>>> requests = new List<Task<IEnumerable<ReaderResultItem>>>(pdfFilesList.Count);
            //for (Int32 n = 0; n < requests.Capacity; n++)
            //    requests.Add(ReaderResultItem.ProcessingPdfFileTask(pdfFilesList[n], tempDir, maxThreads, scanResolution, heightPersentageArea));


            //while (requests.Count > 0)
            //{
            //    // Последовательная обработка каждого завершенного ответа
            //    Task<IEnumerable<ReaderResultItem>> response = await Task.WhenAny(requests);
            //    requests.Remove(response); // Удаление завершенной задачи из коллекции
            //                               // Обработка одного ответа
            //    //Console.WriteLine(response.Result);

            //    var u = response.Result.Reverse();
            //    u.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //}



            //Task taskBarcodeRead = Task.Run(() =>
            //{
            //    //foreach (var pdfFile in pdfFilesList)
            //pdfFilesList.AsParallel().WithDegreeOfParallelism(24).WithMergeOptions(ParallelMergeOptions.FullyBuffered).ForAll(async pdfFile =>
            //{


            //    //////////////////////////////////////////////////////////
            //    ////общая обработка
            //    //////////////////////////////////////////////////////////
            //    //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    //result.Reverse();
            //    //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //    var result = await ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
            //    var u = result.Reverse();
            //    u.ToList().ForEach(i => globalBarcodeItems.Add(i));


            //});
            //});







            //Task.WaitAll(taskBarcodeRead);

            //if (taskBarcodeRead.IsFaulted)
            //{
            //    ConsoleAwesome.WriteLine(taskBarcodeRead.Exception.Message);
            //}
            //else
            //{
            //    ConsoleAwesome.WriteLine(taskBarcodeRead.Status.ToString());
            //}


            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }



        public static void ProcessPDFFolderWithTask(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();



            //List<Task<IEnumerable<ReaderResultItem>>> requests = new List<Task<IEnumerable<ReaderResultItem>>>(pdfFilesList.Count);
            //for (Int32 n = 0; n < requests.Capacity; n++)
            //    requests.Add(ReaderResultItem.ProcessingPdfFileTask(pdfFilesList[n], tempDir, maxThreads, scanResolution, heightPersentageArea));


            //while (requests.Count > 0)
            //{
            //    // Последовательная обработка каждого завершенного ответа
            //    Task<IEnumerable<ReaderResultItem>> response = await Task.WhenAny(requests);
            //    requests.Remove(response); // Удаление завершенной задачи из коллекции
            //                               // Обработка одного ответа
            //    //Console.WriteLine(response.Result);

            //    var u = response.Result.Reverse();
            //    u.ToList().ForEach(i => globalBarcodeItems.Add(i));
            //}



            //Task taskBarcodeRead = Task.Run(() =>
            //{
            //    //foreach (var pdfFile in pdfFilesList)
            pdfFilesList.AsParallel().WithDegreeOfParallelism(24).WithMergeOptions(ParallelMergeOptions.FullyBuffered).ForAll(async pdfFile =>
            {


                //////////////////////////////////////////////////////////
                ////общая обработка
                //////////////////////////////////////////////////////////
                //var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
                //result.Reverse();
                //result.ToList().ForEach(i => globalBarcodeItems.Add(i));
                var result = await ReaderResultItem.ProcessingPdfFileTask(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
                var u = result.Reverse();
                u.ToList().ForEach(i => globalBarcodeItems.Add(i));


            });
            //});







            //Task.WaitAll(taskBarcodeRead);

            //if (taskBarcodeRead.IsFaulted)
            //{
            //    ConsoleAwesome.WriteLine(taskBarcodeRead.Exception.Message);
            //}
            //else
            //{
            //    ConsoleAwesome.WriteLine(taskBarcodeRead.Status.ToString());
            //}


            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }





        public static void ProcessPDFFolderParallelFor(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();


            Parallel.For(0, pdfFilesList.Count, new ParallelOptions { MaxDegreeOfParallelism = -1 }, i =>
            {
                //////////////////////////////////////////////////////////
                ////общая обработка
                //////////////////////////////////////////////////////////

                var result = ReaderResultItem.ProcessingPdfFileTask(pdfFilesList[i], tempDir, maxThreads, scanResolution, heightPersentageArea);
                var u = result.Result.Reverse();
                u.ToList().ForEach(j => globalBarcodeItems.Add(j));
            });





            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }



        public static void ProcessPDFFolderWithSemaphore(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();



            var semaphoreSlim = new SemaphoreSlim(maxThreads);
            var tasks = new List<Task>(pdfFilesList.Count);

            foreach (var pdfFile in pdfFilesList)
            {
                tasks.Add(Task.Run(() =>
                {
                    semaphoreSlim.Wait();
                    try
                    {
                        //////////////////////////////////////////////////////////
                        ////общая обработка
                        var result = ReaderResultItem.ProcessingPdfFile(pdfFile, tempDir, maxThreads, scanResolution, heightPersentageArea);
                        foreach (var res in result) globalBarcodeItems.Add(res);



                    }
                    finally
                    {
                        semaphoreSlim.Release();
                    }
                }));
            }
            Task.WaitAll(tasks.ToArray());


            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }


        public static void ProcessPDFFolderWithSemaphoreAndStreams(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();

            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 


            Stopwatch timerMemoryStream = new Stopwatch();
            timerMemoryStream.Start();
            double totalTimeM = 0;
            double totalSecondsM = 0.0f;
            string totalSecondsValueM = "";

            ConcurrentDictionary<string, MemoryStream> memoryStreams = new ConcurrentDictionary<string, MemoryStream>();
            pdfFilesList.AsParallel().ForAll(pdfFile =>
            {
                var memoryStream = new MemoryStream(File.ReadAllBytes(pdfFile))
                {
                    Position = 0
                };

                memoryStreams.TryAdd(pdfFile, memoryStream);
            });



            timerMemoryStream.Stop();
            totalSecondsM = timerMemoryStream.Elapsed.TotalSeconds;
            totalTimeM += totalSeconds;
            totalSecondsValueM = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На загрузку всех файлов в память затраченно = " + totalSecondsM + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);





            var semaphoreSlim = new SemaphoreSlim(maxThreads);
            var tasks = new List<Task>(pdfFilesList.Count);

            foreach (var dictItem in memoryStreams)
            {
                tasks.Add(Task.Run(() =>
                {
                    semaphoreSlim.Wait();
                    try
                    {
                        //////////////////////////////////////////////////////////
                        ////общая обработка
                        var pdfFile=dictItem.Key;
                        var pdfFileStream = dictItem.Value;


                        var result = ReaderResultItem.ProcessingPdfFile(pdfFile, pdfFileStream, tempDir, maxThreads, scanResolution, heightPersentageArea,false);
                        
                       
                        foreach(var res in result) globalBarcodeItems.Add(res);
                        


                    }
                    finally
                    {
                        semaphoreSlim.Release();
                    }
                }));
            }
            Task.WaitAll(tasks.ToArray());


            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }



        public static async void ProcessPDFFolderWithTaskAlt(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }


            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();


            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 
            ///

            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();

            ThreadPool.SetMinThreads(24, 100);
            ThreadPool.SetMaxThreads(24, 100);

            var t = Enumerable.Range(0, pdfFilesList.Count()).Select(fileIdx => new Func<Task<IEnumerable<ReaderResultItem>>>(() => ReaderResultItem.ProcessingPdfFileTask(pdfFilesList[fileIdx], tempDir, maxThreads, scanResolution, heightPersentageArea)));
            var list = new List<IEnumerable<ReaderResultItem>>();
            await ParallelForEachAsync(t, 24, async func =>
            {
                list.Add(await func());
            });

            //var hj = list.Count();


            list.AsParallel().ForAll(list1 =>
            {

                list1.Reverse();
                list1.ToList().ForEach(i => globalBarcodeItems.Add(i));
            });






            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

        }

        public async static void ProcessPDFFolderWithTaskAltWithStreams(string folderWithPdf, string tempFolderPath, string excellOutputFileNameWithoutExt = "Выгрузка", int maxThreads = 4, int scanResolution = 300, int heightPersentageArea = 10)
        {


            string tempDir = "";
            if (String.IsNullOrEmpty(tempFolderPath))
            {
                tempDir = System.IO.Path.GetTempPath();
            }
            else
            {
                tempDir = tempFolderPath;
            }


            //Удаляем хвосты прошлых запусков
            foreach (var f in Directory.GetDirectories(tempDir, "*PDFProcessingAntonovKA*"))
            {
                Directory.Delete(f, true);
            }

            var pdfFilesList = Directory.GetFiles(folderWithPdf, @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();
            //var pdfFilesList = Directory.GetFiles(@"C:\Users\kosPC\Desktop\работа\Проекты\Фазлеев(Barcodes)\TestsFolder\test100", @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();
            //var pdfFilesList2 = Directory.GetFiles(@"C:\Users\kosPC\Desktop\работа\Проекты\Фазлеев(Barcodes)\TestsFolder\test200", @"*.pdf", SearchOption.AllDirectories).OrderBy(f => f.ToLower()).ToList();

            Stopwatch timerGlobal = new Stopwatch();
            timerGlobal.Start();
            double totalTime = 0;
            double totalSeconds = 0.0f;
            string totalSecondsValue = "";

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("Запускаем таймер для обработки всех документов", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);

            ////////////////////////////////////////////////////////////////////////////////////// 


            Stopwatch timerMemoryStream = new Stopwatch();
            timerMemoryStream.Start();
            double totalTimeM = 0;
            double totalSecondsM = 0.0f;
            string totalSecondsValueM = "";

            ConcurrentDictionary<string, MemoryStream> memoryStreams = new ConcurrentDictionary<string, MemoryStream>();
            pdfFilesList.AsParallel().ForAll(pdfFile =>
            {
                var memoryStream = new MemoryStream(File.ReadAllBytes(pdfFile))
                {
                    Position = 0
                };

                memoryStreams.TryAdd(pdfFile, memoryStream);
            });


            //ConcurrentDictionary<string, MemoryStream> memoryStreams2 = new ConcurrentDictionary<string, MemoryStream>();
            //pdfFilesList2.AsParallel().ForAll(pdfFile => {
            //    var memoryStream = new MemoryStream(File.ReadAllBytes(pdfFile))
            //    {
            //        Position = 0
            //    };

            //    memoryStreams2.TryAdd(pdfFile, memoryStream);
            //});


            timerMemoryStream.Stop();
            totalSecondsM = timerMemoryStream.Elapsed.TotalSeconds;
            totalTimeM += totalSeconds;
            totalSecondsValueM = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На загрузку всех файлов в память затраченно = " + totalSecondsM + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);








            var globalBarcodeItems = new ConcurrentBag<ReaderResultItem>();



            ThreadPool.SetMinThreads(12, 1000);
            ThreadPool.SetMaxThreads(24, 1000);

            var dictArray = memoryStreams.ToArray();
            //var dictArray = memoryStreams.ToList();

            var t = Enumerable.Range(0, dictArray.Count()).Select(fileIdx => new Func<Task<IEnumerable<ReaderResultItem>>>(() => ReaderResultItem.ProcessingPdfFileTask(dictArray[fileIdx].Key, dictArray[fileIdx].Value, tempDir, maxThreads, scanResolution, heightPersentageArea, false)));
            var list = new List<IEnumerable<ReaderResultItem>>();
            await ParallelForEachAsync(t, 12, async func =>
            {
                list.Add(await func());
            });




            list.AsParallel().ForAll(list1 =>
            {

                //list1.Reverse();
                list1.ToList().ForEach(i => globalBarcodeItems.Add(i));
            });


            //dictArray.ForEach(dic =>
            //{
            //    var res = ReaderResultItem.ProcessingPdfFileTask(dic.Key, dic.Value, tempDir, maxThreads, scanResolution, heightPersentageArea, false);
            //    res.Result.ToList().ForEach(i => globalBarcodeItems.Add(i));

            //});


            timerGlobal.Stop();
            totalSeconds = timerGlobal.Elapsed.TotalSeconds;
            totalTime += totalSeconds;
            totalSecondsValue = string.Format("{0:0.00}", totalSeconds);

            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("На обработку всех документов затрачено = " + totalSecondsValue + " sec.", ConsoleColor.Yellow);
            ConsoleAwesome.WriteLine("-------------------------------------", ConsoleColor.Yellow);




            /////////////////////////////////////////////////////////////////////////////////////




            ConsoleAwesome.WriteLine("Формируем выгрузку Ексель");
            if (globalBarcodeItems.Count > 0)
            {


                ReaderResultItem.CreateExcell(folderWithPdf, excellOutputFileNameWithoutExt, globalBarcodeItems);
            }




        }

        public static Task ParallelForEachAsync<T>(IEnumerable<T> source, int degreeOfParallelization, Func<T, Task> body)
        {
            async Task AwaitPartition(IEnumerator<T> partition)
            {
                using (partition)
                {
                    while (partition.MoveNext())
                    {
                        await body(partition.Current);
                    }
                }
            }

            return Task.WhenAll(
                Partitioner.Create(source).GetPartitions(degreeOfParallelization).AsParallel().WithMergeOptions(ParallelMergeOptions.FullyBuffered).Select(AwaitPartition)
                );
        }


        public static Tuple<string, string, string, int, int, int, bool> GetStartParams(string[] args)
        {
            //Путь до папки с файлами документации
            string folderWithPdf = "";

            //Имя ексель файла для сохранения полученных данныых
            string excellOutputFileNameWithoutExt = "";

            //временная директория
            string tempFolderPath = "";

            //количество потоков
            string maxThreads = "";
            int maxThreadsParced = 0;


            string scanResolution = "";
            int scanResolutionParced = 0;

            //Область в которой ищем баркод на листе вычислиятся по высоте указанных в процентах
            string heightPersentageArea = "";
            int heightPersentageAreaParced = 0;
            string useMemoryToLoadAllFiles = "";
            bool useMemoryToLoadAllFilesParced = true;
            bool showHelp = false;

            OptionSet options = new OptionSet()
            {
                {"f=|folderWithPDF=", "Путь до папки с файлами документации", v => {folderWithPdf=v.Replace("\n","").Replace("\r",""); } },
                {"e=|excellOutputFileNameWithoutExt=", "Имя файла выгрузки", v => {excellOutputFileNameWithoutExt=v.Replace("\n","").Replace("\r",""); } },
                {"t=|tempFolderPath=", "Путь к временной папке", v => {tempFolderPath=v.Replace("\n","").Replace("\r",""); } },
                {"m=|maxThreads=", "Количество потоков", v => {maxThreads=v.Replace("\n","").Replace("\r",""); } },
                {"r=|scanResolution=", "Разрешение скана", v => {scanResolution=v.Replace("\n","").Replace("\r",""); } },
                {"p=|heightPersentageArea=", "Область поиска баркода", v => {heightPersentageArea=v.Replace("\n","").Replace("\r",""); } },
                {"u=|useMemoryToLoadAllFiles=", "Хранить файлы в памяти", v => {useMemoryToLoadAllFiles=v.Replace("\n","").Replace("\r",""); } },
                {"h|help|?", "Описание всех консольныых аргументов приложения", v => {showHelp = v != null; } }
            };
            try
            {
                //все что не распозналось
                var extra = options.Parse(args);

                if ((!args.Any() && extra.Any()) || extra.Any())
                {
                    showHelp = true;
                }

                if (showHelp)
                {
                    //выводим хелп
                    options.WriteOptionDescriptions(Console.Out);
                    //return;
                }

            }
            catch
            {
                Console.WriteLine("Ошибка получения значений аргументов консоли");
                throw new Exception("Ошибка получения значений аргументов консоли");
            };

            //folderWithPdf = GetArgument(args, "folderWithPdf");

            //excellTemplate = GetArgument(args, "excellTemplate");

            //excellOutputFileNameWithoutExt = GetArgument(args, "excellOutputFileNameWithoutExt");

            Console.WriteLine("-------------------------------------");
            Console.WriteLine("Парамметры запуска приложения: ");


            if (Directory.Exists(folderWithPdf))
            {
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Путь до папки с файлами документации: {0} \nполучен из аргументов консоли", folderWithPdf));
            }
            else
            {

                folderWithPdf = ReadSetting("FolderWithPdf");
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Путь до папки с файлами документации: {0} \nполучен из файла App.config", folderWithPdf));
            }


            if (!string.IsNullOrEmpty(excellOutputFileNameWithoutExt))
            {
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Имя ексель файла для сохранения полученных данных: {0} \nполучено из аргументов консоли", excellOutputFileNameWithoutExt));
            }
            else
            {
                excellOutputFileNameWithoutExt = ReadSetting("ExcellOutputFileNameWithoutExt");
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Имя ексель файла для сохранения полученных данных: {0} \nполучено из файла App.config", excellOutputFileNameWithoutExt));
            }


            if (!string.IsNullOrEmpty(tempFolderPath))
            {
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Путь до временной папки: {0} \nполучено из аргументов консоли", tempFolderPath));
            }
            else
            {
                tempFolderPath = ReadSetting("TempDirPath");
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Путь до временной папки: {0} \nполучено из файла App.config", tempFolderPath));
            }

            if (!string.IsNullOrEmpty(maxThreads))
            {
                int m;
                if (int.TryParse(maxThreads, out m))
                {
                    maxThreadsParced = m;
                }
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Количество потоков для параллельного выполненния: {0} \nполучено из аргументов консоли", maxThreads));
            }
            else
            {
                int m;
                if (int.TryParse(ReadSetting("MaxThreads"), out m))
                {
                    maxThreadsParced = m;
                }

                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Количество потоков для параллельного выполненния: {0} \nполучено из файла App.config", maxThreadsParced));
            }

            if (!string.IsNullOrEmpty(scanResolution))
            {
                int m;
                if (int.TryParse(scanResolution, out m))
                {
                    scanResolutionParced = m;
                }
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Разрешение скана: {0} \nполучено из аргументов консоли", scanResolutionParced));
            }
            else
            {
                int m;
                if (int.TryParse(ReadSetting("ScanResolution"), out m))
                {
                    scanResolutionParced = m;
                }

                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Разрешение скана: {0} \nполучено из файла App.config", scanResolutionParced));
            }


            if (!string.IsNullOrEmpty(heightPersentageArea))
            {
                int m;
                if (int.TryParse(heightPersentageArea, out m))
                {
                    heightPersentageAreaParced = m;
                }
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Область поиска баркода в %: {0} \nполучено из аргументов консоли", heightPersentageAreaParced));
            }
            else
            {
                int m;
                if (int.TryParse(ReadSetting("HeightPersentageArea"), out m))
                {
                    heightPersentageAreaParced = m;
                }

                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Область поиска баркода в %: {0} \nполучено из файла App.config", heightPersentageAreaParced));
            }

            if (!string.IsNullOrEmpty(useMemoryToLoadAllFiles))
            {
                
                if (useMemoryToLoadAllFiles=="true"|| useMemoryToLoadAllFiles == "True")
                {
                    useMemoryToLoadAllFilesParced = true;
                }
                else if (useMemoryToLoadAllFiles == "false" || useMemoryToLoadAllFiles == "False")
                {
                    useMemoryToLoadAllFilesParced = false;
                }
                else
                {
                    useMemoryToLoadAllFilesParced = true;
                }

                
                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Область поиска баркода в %: {0} \nполучено из аргументов консоли", useMemoryToLoadAllFilesParced));
            }
            else
            {
                var boolVal = ReadSetting("UseMemoryToLoadAllFiles");
                if (boolVal == "true" || boolVal == "True")
                {
                    useMemoryToLoadAllFilesParced = true;
                }
                else if (boolVal == "false" || boolVal == "False")
                {
                    useMemoryToLoadAllFilesParced = false;
                }
                else
                {
                    useMemoryToLoadAllFilesParced = true;
                }

                Console.WriteLine("-------------------------------------");
                Console.WriteLine(String.Format("Область поиска баркода в %: {0} \nполучено из файла App.config", useMemoryToLoadAllFilesParced));
            }

            Console.WriteLine("-------------------------------------");

            return new Tuple<string, string, string, int, int, int,bool>(folderWithPdf, tempFolderPath, excellOutputFileNameWithoutExt, maxThreadsParced, scanResolutionParced, heightPersentageAreaParced, useMemoryToLoadAllFilesParced);

        }

        static string ReadSetting(string key)
        {
            string result = "";
            try
            {
                var appSettings = ConfigurationManager.AppSettings;

                result = appSettings[key];
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                Console.WriteLine("Ошибка чтения файла конфигурации");
            }
            return result;
        }

    }
}
