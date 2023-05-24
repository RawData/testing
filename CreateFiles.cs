using System;
using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OfficeFilesCreator
{
    class Program
    {
        static List<string> createdFiles = new List<string>();

        static void Main(string[] args)
        {
            string directoryPath = "C:\\Path\\To\\Your\\Directory";
            foreach (var file in Directory.GetFiles(directoryPath))
            {
                // Check if the file name starts with '~' or '$'
                if (Path.GetFileName(file).StartsWith("~") || Path.GetFileName(file).StartsWith("$"))
                {
                    continue; // Skip this file and move to the next one
                }

                string fileName = Path.GetFileName(file);
                Console.WriteLine($"Embedding file: {fileName}");

                CreateWordDocument(file);
                CreateExcelWorkbook(file);
                CreatePowerPointPresentation(file);
                CreateWordTemplate(file);
                CreateExcelTemplate(file);
                CreatePowerPointTemplate(file);
                CreateExcelAddIn(file);
                CreatePowerPointShow(file);
            }

            PrintCreatedFiles();

            Console.WriteLine("Files created successfully!");
        }

        // All the other methods are similar to the CreateWordDocument method
        // They differ in the file format they save in and the type of Microsoft Office application they use

        static void CreateWordDocument(string fileToEmbed)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();

            doc.Content.Text = "Hello, Word Document!";
            doc.InlineShapes.AddOLEObject(ClassType: "Package", FileName: fileToEmbed, LinkToFile: false, DisplayAsIcon: false);
            string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.doc";
            string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
            doc.SaveAs2(newFilePath, WdSaveFormat.wdFormatDocument);
            doc.Close();

            wordApp.Quit();

            createdFiles.Add(newFilePath);
        }

        static void CreateExcelWorkbook(string fileToEmbed)
        {
            var excelApp = new Excel.Application();
            excelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            var workbook = excelApp.Workbooks.Add();

            if (workbook.Sheets.Count > 0)
            {
                var sheet = workbook.Sheets[1] as Excel.Worksheet;
                sheet.Cells[1, 1].Value = "Hello, Excel Workbook!";
                var range = sheet.Range[sheet.Cells[5, 5], sheet.Cells[5, 5]];
                sheet.OLEObjects().Add(Filename: fileToEmbed, Link: MsoTriState.msoFalse, DisplayAsIcon: MsoTriState.msoFalse, Left: range.Left, Top: range.Top, Width: range.Width, Height: range.Height);

                string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.xls";
                string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
                workbook.SaveAs(newFilePath, Excel.XlFileFormat.xlWorkbookDefault);

                createdFiles.Add(newFilePath);
            }

            workbook.Close();
            excelApp.Quit();
        }

        static void CreatePowerPointPresentation(string fileToEmbed)
        {
            var powerpointApp = new PowerPoint.Application();
            var presentation = powerpointApp.Presentations.Add(MsoTriState.msoTrue);

            PowerPoint.Slide slide;
            if (presentation.Slides.Count < 1)
            {
                slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            }
            else
            {
                slide = presentation.Slides[1];
            }

            float left = (slide.Master.Width - 100) / 2;
            float top = (slide.Master.Height - 100) / 2;
            slide.OLEObjects().Add(FileName: fileToEmbed, Link: MsoTriState.msoFalse, DisplayAsIcon: MsoTriState.msoFalse, Left: left, Top: top, Width: 100, Height: 100);

            string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.ppt";
            string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
            presentation.SaveAs(newFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsPresentation);

            createdFiles.Add(newFilePath);

            presentation.Close();
            powerpointApp.Quit();
        }

        static void CreateWordTemplate(string fileToEmbed)
        {
            var wordApp = new Application();
            var doc = wordApp.Documents.Add();

            doc.Content.Text = "Hello, Word Template!";
            doc.InlineShapes.AddOLEObject(ClassType: "Package", FileName: fileToEmbed, LinkToFile: false, DisplayAsIcon: false);

            string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.dotx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
            doc.SaveAs2(newFilePath, WdSaveFormat.wdFormatXMLTemplate);
            doc.Close();

            wordApp.Quit();

            createdFiles.Add(newFilePath);
        }

        static void CreateExcelTemplate(string fileToEmbed)
        {
            var excelApp = new Excel.Application();
            excelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            var workbook = excelApp.Workbooks.Add();

            if (workbook.Sheets.Count > 0)
            {
                var sheet = workbook.Sheets[1] as Excel.Worksheet;
                sheet.Cells[1, 1].Value = "Hello, Excel Template!";
                var range = sheet.Range[sheet.Cells[5, 5], sheet.Cells[5, 5]];
                sheet.OLEObjects().Add(Filename: fileToEmbed, Link: MsoTriState.msoFalse, DisplayAsIcon: MsoTriState.msoFalse, Left: range.Left, Top: range.Top, Width: range.Width, Height: range.Height);

                string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.xltx";
                string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
                workbook.SaveAs(newFilePath, Excel.XlFileFormat.xlOpenXMLTemplate);

                createdFiles.Add(newFilePath);
            }

            workbook.Close();
            excelApp.Quit();
        }

        static void CreatePowerPointTemplate(string fileToEmbed)
        {
            var powerpointApp = new PowerPoint.Application();
            var presentation = powerpointApp.Presentations.Add(MsoTriState.msoTrue);

            PowerPoint.Slide slide;
            if (presentation.Slides.Count < 1)
            {
                slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            }
            else
            {
                slide = presentation.Slides[1];
            }

            float left = (slide.Master.Width - 100) / 2;
            float top = (slide.Master.Height - 100) / 2;
            slide.OLEObjects().Add(FileName: fileToEmbed, Link: MsoTriState.msoFalse, DisplayAsIcon: MsoTriState.msoFalse, Left: left, Top: top, Width: 100, Height: 100);

            string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.potx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
            presentation.SaveAs(newFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsTemplate);

            createdFiles.Add(newFilePath);

            presentation.Close();
            powerpointApp.Quit();
        }

        static void CreateExcelAddIn(string fileToEmbed)
        {
            var excelApp = new Excel.Application();
            excelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            var workbook = excelApp.Workbooks.Add();

            if (workbook.Sheets.Count > 0)
            {
                var sheet = workbook.Sheets[1] as Excel.Worksheet;
                sheet.Cells[1, 1].Value = "Hello, Excel Add-In!";
                var range = sheet.Range[sheet.Cells[5, 5], sheet.Cells[5, 5]];
                sheet.OLEObjects().Add(Filename: fileToEmbed, Link: MsoTriState.msoFalse, DisplayAsIcon: MsoTriState.msoFalse, Left: range.Left, Top: range.Top, Width: range.Width, Height: range.Height);

                string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.xlam";
                string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
                workbook.SaveAs(newFilePath, Excel.XlFileFormat.xlOpenXMLAddIn);

                createdFiles.Add(newFilePath);
            }

            workbook.Close();
            excelApp.Quit();
        }

        static void CreatePowerPointShow(string fileToEmbed)
        {
            var powerpointApp = new PowerPoint.Application();
            var presentation = powerpointApp.Presentations.Add(MsoTriState.msoTrue);

            PowerPoint.Slide slide;
            if (presentation.Slides.Count < 1)
            {
                slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            }
            else
            {
                slide = presentation.Slides[1];
            }

            float left = (slide.Master.Width - 100) / 2;
            float top = (slide.Master.Height - 100) / 2;
            slide.OLEObjects().Add(FileName: fileToEmbed, Link: MsoTriState.msoFalse, DisplayAsIcon: MsoTriState.msoFalse, Left: left, Top: top, Width: 100, Height: 100);

            string newFileName = Path.GetFileNameWithoutExtension(fileToEmbed) + "_DSAS.ppsx";
            string newFilePath = Path.Combine(Path.GetDirectoryName(fileToEmbed), newFileName);
            presentation.SaveAs(newFilePath, PowerPoint.PpSaveAsFileType.ppSaveAsShow);

            createdFiles.Add(newFilePath);

            presentation.Close();
            powerpointApp.Quit();
        }

        static void PrintCreatedFiles()
        {
            Console.WriteLine("Created files:");
            foreach (var file in createdFiles)
            {
                Console.WriteLine(file);
            }
        }
    }
}
