using TesseractOCR.Enums;
using TesseractOCR;
using System.Configuration;
using PDFtoImage;
using System;
using System.Drawing;
using AForge.Imaging.Filters;
using AForge.Imaging;
using SkiaSharp;

namespace DSC
{
    /// <summary>
    /// Program that reads text from images and files and interprets that into correct and summarized information.
    /// Author: Pedro Sanches
    /// Version: 0.02 - Added PreProcessing option
    /// </summary>
    class Program
    {
        static async Task Main()
        {
            Console.WriteLine("******************************");
            await ReadPDF();
            Console.WriteLine("******************************");
            Console.ReadKey();
        }

        /// <summary>
        /// Method that reads data from each invoice in a specific path
        /// </summary>
        /// <returns>Nothing</returns>
        public static async Task ReadPDF()
        {
            string pathFiles = AppDomain.CurrentDomain.BaseDirectory + System.Configuration.ConfigurationManager.AppSettings.Get("PathToFiles");
            string pathTessData = AppDomain.CurrentDomain.BaseDirectory + "tessdata";
            string[] files = Directory.GetFiles(pathFiles, "*.pdf");

            bool addPreProcessing = false;
            bool.TryParse(System.Configuration.ConfigurationManager.AppSettings.Get("AddPreProcessing"), out addPreProcessing);


            foreach (string file in files)
            {

                Console.WriteLine("Name of the file: " + System.IO.Path.GetFileNameWithoutExtension(file));

                var pagesFromFile = await GetImage(file, addPreProcessing);

                int pageCount = 0;

                //Check on each page from this file
                foreach (var load in pagesFromFile)
                {
                    pageCount++;

                    var engine = new Engine(pathTessData, "eng", EngineMode.TesseractAndLstm);

                    var read = await GetPage(load, engine);

                    switch (read.MeanConfidence)
                    {
                        case float n when n < 0.6:
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Page {pageCount}/{pagesFromFile.Count} - Quality: Low({n})");
                            goto default;
                        case float n when n >= 0.6 && n < 0.65:
                            Console.ForegroundColor = ConsoleColor.DarkYellow;
                            Console.WriteLine($"Page {pageCount}/{pagesFromFile.Count} - Quality: Medium ({n})");
                            goto default;
                        case float n when n >= 0.65 && n < 0.8:
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Page {pageCount}/{pagesFromFile.Count} - Quality: Good ({n})");
                            goto default;
                        case float n when n >= 0.8 && n < 0.9:
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"Page {pageCount}/{pagesFromFile.Count} - Quality: Very good ({n})");
                            goto default;
                        case float n when n >= 0.9:
                            Console.ForegroundColor = ConsoleColor.DarkGreen;
                            Console.WriteLine($"Page {pageCount}/{pagesFromFile.Count} - Quality: Excellent ({n})");
                            goto default;
                        default:
                            Console.ForegroundColor = ConsoleColor.White;
                            break;
                    }

                    try
                    {
                        using (StreamWriter writer = File.CreateText($"{pathFiles}\\{System.IO.Path.GetFileNameWithoutExtension(file)}-Page{pageCount}.txt"))
                        {
                            writer.WriteLine(read.Text);
                        }

                    }
                    catch (Exception ex)
                    {

                    }

                    load.Dispose();
                    engine.Dispose();
                }

            }

        }

        public static async Task<List<TesseractOCR.Pix.Image>> GetImage(string file, bool addPreProcessing)
        {

            List<TesseractOCR.Pix.Image> images = new();
            List<MemoryStream> memoryStreamsList = await GetStreams(file);

            //Control flag - Check if preproccessing is possible
            bool state = false;

            foreach (var ms in memoryStreamsList)
            {

                //Only if PreProcessing is true in the App configuration
                if (addPreProcessing)
                {

                    //Convert to Bitmap
                    try
                    {
                        
#pragma warning disable CA1416 // Validate platform compatibility

                        var bitmap = new Bitmap(ms); //Only supported on Windows

#pragma warning restore CA1416 // Validate platform compatibility

                        TesseractOCR.Pix.Image img = await PreprocessImage(bitmap);
                        images.Add(img);
                        state = true;

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error when trying to PreProccess Image: " + ex.Message + " - " + ex.StackTrace);
                    }

                }

                //Only if preproccess didn't work, we just load the image normally
                if (!state)
                {
                    await Task.Run(() => images.Add(TesseractOCR.Pix.Image.LoadFromMemory(ms)));
                }
                state = false;

            }

            return images;

        }


        public static async Task<TesseractOCR.Pix.Image> PreprocessImage(Bitmap image)
        {

            // Resize and scale
            var resizedImage = new Bitmap(image, new Size(image.Width * 2, image.Height * 2)); // Double the size for better resolution

            // Convert to grayscale
            var grayscaleImage = Grayscale.CommonAlgorithms.BT709.Apply(resizedImage);

            // Apply binarization
            var thresholdFilter = new BradleyLocalThresholding();
            var binarizedImage = thresholdFilter.Apply(grayscaleImage);

            // Enhance contrast
            var contrastFilter = new ContrastStretch();
            contrastFilter.ApplyInPlace(binarizedImage);

            //Convert from Bitmap to MemoryStream
            using (MemoryStream memoryStream = new MemoryStream())
            {

                // Save the bitmap to the MemoryStream
                grayscaleImage.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Bmp);

                return await Task.Run(() => TesseractOCR.Pix.Image.LoadFromMemory(memoryStream));

            }

        }

        public static async Task<List<MemoryStream>> GetStreams(string file)
        {
            //Console.WriteLine(Path.GetFileNameWithoutExtension(file));
            var pdf = File.ReadAllBytes(file);

#pragma warning disable CA1416 // Validate platform compatibility

            var convertions = Conversion.ToImages(pdf);

#pragma warning restore CA1416 // Validate platform compatibility

            List<MemoryStream> streams = new List<MemoryStream>();

            foreach (var convert in convertions)
            {
                var output = convert.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100).AsStream();
                MemoryStream ms = new MemoryStream();

                output.CopyTo(ms);
                await Task.Run(() => streams.Add(ms));
            }

            return streams;

        }

        public static async Task<TesseractOCR.Page> GetPage(TesseractOCR.Pix.Image image, TesseractOCR.Engine engine)
        {
            return await Task.Run(() => engine.Process(image));
        }

    }
}