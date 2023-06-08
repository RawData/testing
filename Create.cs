using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;

namespace OfficeFilesCreator
{
    class Program
    {
        static List<string> createdFiles = new List<string>();

        static void Main(string[] args)
        {
            string directoryPath = @"C:\Path\To\Your\Directory";
            foreach (var file in Directory.GetFiles(directoryPath))
            {
              
                if (Path.GetFileName(file).StartsWith("~") || Path.GetFileName(file).StartsWith("$"))
                {
                    continue; // Skip this file and move to the next one
                }

                string fileName = Path.GetFileName(file);
                Console.WriteLine("Embedding file: " + fileName);

                string newFilePath = EmbedFilesInPresentation(file);

                createdFiles.Add(newFilePath);
            }

            PrintCreatedFiles();

            Console.WriteLine("Files created successfully!");
            Console.ReadLine();
        }

        static string EmbedFilesInPresentation(string fileToEmbed)
        {
            var powerPointApp = new Application();
            var presentation = powerPointApp.Presentations.Add(MsoTriState.msoTrue);

            Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            float left = 0;
            float top = 0;
            float width = 500;
            float height = 400;

            Shape oleShape = slide.Shapes.AddOLEObject(
                Left: left,
                Top: top,
                Width: width,
                Height: height,
                ClassName: "Package",
                FileName: fileToEmbed,
                LinkToFile: MsoTriState.msoFalse,
                DisplayAsIcon: MsoTriState.msoFalse
            );

            string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_Embedded.ppt";
            string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
            presentation.SaveAs(newFilePath, PpSaveAsFileType.ppSaveAsDefault);

            presentation.Close();
            powerPointApp.Quit();

            return newFilePath;
        }

        static void PrintCreatedFiles()
        {
            Console.WriteLine("List of Created Files:");
            foreach (var file in createdFiles)
            {
                Console.WriteLine(file);
            }
        }

        static void CreateWordDocument(string filePath)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();
            doc.Content.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".doc";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            doc.SaveAs(newFilePath, WdSaveFormat.wdFormatDocument);

            doc.Close();
            wordApp.Quit();
        }

        static void CreateWordMacroEnabledDocument(string filePath)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();
            doc.Content.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".docm";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            doc.SaveAs(newFilePath, WdSaveFormat.wdFormatXMLDocumentMacroEnabled);

            doc.Close();
            wordApp.Quit();
        }

        static void CreateWordDocumentOpenXml(string filePath)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();
            doc.Content.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".docx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            doc.SaveAs(newFilePath, WdSaveFormat.wdFormatXMLDocument);

            doc.Close();
            wordApp.Quit();
        }

        static void CreateWordTemplate(string filePath)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();
            doc.Content.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".dot";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            doc.SaveAs(newFilePath, WdSaveFormat.wdFormatTemplate);

            doc.Close();
            wordApp.Quit();
        }

        static void CreateWordMacroEnabledTemplate(string filePath)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();
            doc.Content.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".dotm";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            doc.SaveAs(newFilePath, WdSaveFormat.wdFormatXMLTemplateMacroEnabled);

            doc.Close();
            wordApp.Quit();
        }

        static void CreateWordTemplateOpenXml(string filePath)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();
            doc.Content.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".dotx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            doc.SaveAs(newFilePath, WdSaveFormat.wdFormatXMLTemplate);

            doc.Close();
            wordApp.Quit();
        }

        static void CreatePowerPointSlideshow(string filePath)
        {
            var powerPointApp = new Application();
            var presentation = powerPointApp.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank).Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 400, 200).TextFrame.TextRange.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".pps";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            presentation.SaveAs(newFilePath, PpSaveAsFileType.ppSaveAsDefault);

            presentation.Close();
            powerPointApp.Quit();
        }

        static void CreatePowerPointMacroEnabledSlideshow(string filePath)
        {
            var powerPointApp = new Application();
            var presentation = powerPointApp.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank).Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 400, 200).TextFrame.TextRange.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".ppsm";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            presentation.SaveAs(newFilePath, PpSaveAsFileType.ppSaveAsDefault);

            presentation.Close();
            powerPointApp.Quit();
        }

        static void CreatePowerPointSlideshowOpenXml(string filePath)
        {
            var powerPointApp = new Application();
            var presentation = powerPointApp.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank).Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 400, 200).TextFrame.TextRange.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".ppsx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            presentation.SaveAs(newFilePath, PpSaveAsFileType.ppSaveAsDefault);

            presentation.Close();
            powerPointApp.Quit();
        }

        static void CreatePowerPointPresentation(string filePath)
        {
            var powerPointApp = new Application();
            var presentation = powerPointApp.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank).Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 400, 200).TextFrame.TextRange.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".ppt";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            presentation.SaveAs(newFilePath, PpSaveAsFileType.ppSaveAsDefault);

            presentation.Close();
            powerPointApp.Quit();
        }

        static void CreatePowerPointMacroEnabledPresentation(string filePath)
        {
            var powerPointApp = new Application();
            var presentation = powerPointApp.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank).Shapes.AddTextbox(
                MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 400, 200).TextFrame.TextRange.Text = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".pptm";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            presentation.SaveAs(newFilePath, PpSaveAsFileType.ppSaveAsDefault);

            presentation.Close();
            powerPointApp.Quit();
        }

        static void CreateExcelWorkbook(string filePath)
        {
            var excelApp = new Application();
            var workbook = excelApp.Workbooks.Add();
            var worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Cells[1, 1] = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".xls";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            workbook.SaveAs(newFilePath, XlFileFormat.xlWorkbookNormal);

            workbook.Close();
            excelApp.Quit();
        }

        static void CreateExcelMacroEnabledWorkbook(string filePath)
        {
            var excelApp = new Application();
            var workbook = excelApp.Workbooks.Add();
            var worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Cells[1, 1] = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".xlsm";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFileName);
            workbook.SaveAs(newFilePath, XlFileFormat.xlOpenXMLWorkbookMacroEnabled);

            workbook.Close();
            excelApp.Quit();
        }

        static void CreateExcelWorkbookOpenXml(string filePath)
        {
            var excelApp = new Application();
            var workbook = excelApp.Workbooks.Add();
            var worksheet = workbook.Worksheets[1] as Worksheet;
            worksheet.Cells[1, 1] = "Testing";
            string newFileName = Path.GetFileNameWithoutExtension(filePath) + ".xlsx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(filePath), newFilePath);
            workbook.SaveAs(newFilePath, XlFileFormat.xlOpenXMLWorkbook);

            workbook.Close();
            excelApp.Quit();
        }
    }
}
