using ComPDF_Conversion.Common;
using ComPDF_Conversion.Conversion;
using ComPDF_Conversion.DocumentAI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Linq;
using MessageBox = System.Windows.MessageBox;

namespace ComPDF_Conversion_Demo
{
  public class ConvertOptions
  {
    public OCRLanguage OCRLanguage = OCRLanguage.e_ENGLISH;
    public bool ContainAnnotations = true;
    public bool CsvFormat = false;
    public bool AllContent = false;
    public bool OneTablePerSheet = true;
    public bool ContainImages = true;
    public bool ContainPageBackgroundImage = true;
    public bool AutoCreateFolder = false;
    public bool OutputDocumentPerPage = false;
    public bool EnableAiLayout = true;
    public bool EnableOCR = false;
    public bool TxtTableFormat = true;
    public bool FormulaToImage = false;
    public bool ImagePathEnhance = true;
    public bool ContainTables = true;
    public float ImageRatio = 1.0f;
    public PageLayoutMode LayoutMode = PageLayoutMode.e_Flow;
    public ExcelWorksheetOption WorksheetOption = ExcelWorksheetOption.e_ForTable;
    public HtmlPageOption htmlOption = HtmlPageOption.e_SinglePage;
    public ImageType ImageFormat = ImageType.PNG;
    public ImageColorMode ImageMode = ImageColorMode.Color;
    public OCROption OcrOption = OCROption.e_All;
  }

  /// <summary>
  /// Copyright © 2014-2025 PDF Technologies, Inc. All Rights Reserved.
  ///
  /// THIS SOURCE CODE AND ANY ACCOMPANYING DOCUMENTATION ARE PROTECTED BY INTERNATIONAL COPYRIGHT LAW
  /// AND MAY NOT BE RESOLD OR REDISTRIBUTED.USAGE IS BOUND TO THE ComPDFKit LICENSE AGREEMENT.
  /// UNAUTHORIZED REPRODUCTION OR DISTRIBUTION IS SUBJECT TO CIVIL AND CRIMINAL PENALTIES.
  /// This notice may not be removed from this file.
  ///
  /// https://www.compdf.com
  /// </summary>
  public partial class MainWindow : Window
  {
    private OnProgress getPorgress = null;
    private ErrorCode err;
    public ConvertOptions Options;
    private List<string> selectedFiles;
    public MainWindow()
    {
      InitializeComponent();
      getPorgress = GetProgress;
      Options = new ConvertOptions();
    }

    #region Method
    private void GetProgress(int pageIndex, int total)
    {
      Dispatcher.Invoke(() =>
      {
        Progress.Text = pageIndex + "/" + total;
        if (pageIndex == total)
        {
          Progress.Text += " Conversion to complete.";
        }
      });
    }

    private void ShowErrorMessage(ErrorCode error)
    {
      string errorMessage = "";

      switch (error)
      {
        case ErrorCode.e_ErrPDFPassword:
          errorMessage = "Password required or incorrect password.";
          break;

        case ErrorCode.e_ErrLicensePermissionDeny:
          errorMessage = "The license doesn't allow the permission.";
          break;

        case ErrorCode.e_ErrOutOfMemory:
          errorMessage = "Malloc failure.";
          break;

        case ErrorCode.e_ErrFile:
          errorMessage = "File not found or could not be opened.";
          break;

        case ErrorCode.e_ErrPDFFormat:
          errorMessage = "File not in PDF format or corrupted.";
          break;

        case ErrorCode.e_ErrPDFSecurity:
          errorMessage = "Unsupported security scheme.";
          break;

        case ErrorCode.e_ErrPDFPage:
          errorMessage = "Page not found or content error.";
          break;

        case ErrorCode.e_ErrCancel:
          errorMessage = "Conversion task Canceled.";
          break;

        case ErrorCode.e_ErrNoTable:
          errorMessage = "There are no tables in the PDF document.";
          break;

        case ErrorCode.e_ErrConverting:
          errorMessage = "A file is already being converted, so other files could not be converted at the same time.";
          break;

        case ErrorCode.e_ErrUnknown:
          errorMessage = "Unknown error in processing conversion.";
          break;

        default:
          errorMessage = "Unknown error in processing conversion.";
          break;
      }

      MessageBox.Show("Conversion failure : " + errorMessage);
    }

    private async Task WordConvert()
    {
      try
      {
        WordOptions wordOptions = new WordOptions();
        wordOptions.ContainImage = Options.ContainImages;
        wordOptions.ContainAnnotation = Options.ContainAnnotations;
        wordOptions.FormulaToImage = Options.FormulaToImage;
        wordOptions.EnableAiLayout = Options.EnableAiLayout;
        wordOptions.EnableOCR = Options.EnableOCR;
        wordOptions.LayoutMode = Options.LayoutMode;
        wordOptions.ContainPageBackgroundImage = Options.ContainPageBackgroundImage;
        wordOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        wordOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToWord(input, "", outputFolder, wordOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task ExcelConvert()
    {
      try
      {
        ExcelOptions excelOptions = new ExcelOptions();
        excelOptions.ContainImage = Options.ContainImages;
        excelOptions.ContainAnnotation = Options.ContainAnnotations;
        excelOptions.FormulaToImage = Options.FormulaToImage;
        excelOptions.AllContent = Options.AllContent;
        excelOptions.CsvFormat = Options.CsvFormat;
        excelOptions.EnableAiLayout = Options.EnableAiLayout;
        excelOptions.EnableOCR = Options.EnableOCR;
        excelOptions.WorksheetOption = Options.WorksheetOption;
        excelOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        excelOptions.AutoCreateFolder = Options.AutoCreateFolder;
        excelOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToExcel(input, "", outputFolder, excelOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task PptConvert()
    {
      try
      {
        PptOptions pptOptions = new PptOptions();
        pptOptions.ContainImage = Options.ContainImages;
        pptOptions.ContainAnnotation = Options.ContainAnnotations;
        pptOptions.FormulaToImage = Options.FormulaToImage;
        pptOptions.EnableAiLayout = Options.EnableAiLayout;
        pptOptions.EnableOCR = Options.EnableOCR;
        pptOptions.ContainPageBackgroundImage = Options.ContainPageBackgroundImage;
        pptOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        pptOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToPpt(input, "", outputFolder, pptOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task HtmlConvert()
    {
      try
      {
        HtmlOptions htmlOptions = new HtmlOptions();
        htmlOptions.ContainImage = Options.ContainImages;
        htmlOptions.ContainAnnotation = Options.ContainAnnotations;
        htmlOptions.FormulaToImage = Options.FormulaToImage;
        htmlOptions.EnableAiLayout = Options.EnableAiLayout;
        htmlOptions.EnableOCR = Options.EnableOCR;
        htmlOptions.LayoutMode = Options.LayoutMode;
        htmlOptions.HtmlOption = Options.htmlOption;
        htmlOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        htmlOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToHtml(input, "", outputFolder, htmlOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task RtfConvert()
    {
      try
      {
        RtfOptions rtfOptions = new RtfOptions();
        rtfOptions.ContainImage = Options.ContainImages;
        rtfOptions.ContainAnnotation = Options.ContainAnnotations;
        rtfOptions.FormulaToImage = Options.FormulaToImage;
        rtfOptions.EnableAiLayout = Options.EnableAiLayout;
        rtfOptions.EnableOCR = Options.EnableOCR;
        rtfOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        rtfOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToRtf(input, "", outputFolder, rtfOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task PdfConvert()
    {
      try
      {
        SearchablePdfOptions pdfOptions = new SearchablePdfOptions();
        pdfOptions.ContainImage = Options.ContainImages;
        pdfOptions.ContainPageBackgroundImage = Options.ContainPageBackgroundImage;
        pdfOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        pdfOptions.OcrOption = OCROption.e_All;
        pdfOptions.EnableOCR = true;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToSearchablePDF(input, "", outputFolder, pdfOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task TxtConvert()
    {
      try
      {
        TxtOptions txtOptions = new TxtOptions();
        txtOptions.TableFormat = Options.TxtTableFormat;
        txtOptions.EnableAiLayout = Options.EnableAiLayout;
        txtOptions.EnableOCR = Options.EnableOCR;
        txtOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        txtOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToTxt(input, "", outputFolder, txtOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task ImageConvert()
    {
      try
      {
        ImageOptions imageOptions = new ImageOptions();
        imageOptions.ImageScaling = Options.ImageRatio;
        imageOptions.PathEnhance = Options.ImagePathEnhance;
        imageOptions.ImageType = Options.ImageFormat;
        imageOptions.ImageColorMode = Options.ImageMode;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToImage(input, "", Path.Combine(outputFolder, outputFileName), imageOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task JsonConvert()
    {
      try
      {
        JsonOptions jsonOptions = new JsonOptions();
        jsonOptions.ContainImage = Options.ContainImages;
        jsonOptions.ContainTable = Options.ContainTables;
        jsonOptions.EnableAiLayout = Options.EnableAiLayout;
        jsonOptions.EnableOCR = Options.EnableOCR;
        jsonOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        jsonOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToJson(input, "", outputFolder, jsonOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }

    private async Task MarkdownConvert()
    {
      try
      {
        MarkdownOptions markdownOptions = new MarkdownOptions();
        markdownOptions.ContainImage = Options.ContainImages;
        markdownOptions.ContainAnnotation = Options.ContainAnnotations;
        markdownOptions.EnableAiLayout = Options.EnableAiLayout;
        markdownOptions.EnableOCR = Options.EnableOCR;
        markdownOptions.OutputDocumentPerPage = Options.OutputDocumentPerPage;
        markdownOptions.OcrOption = Options.OcrOption;

        string outputFolder = OutputPath.Text;
        string outputFileName = Path.GetFileNameWithoutExtension(InputPath.Text);
        string input = InputPath.Text;
        err = await Task.Run(() => CPDFConversion.StartPDFToMarkdown(input, "", outputFolder, markdownOptions));
        if (err != ErrorCode.e_ErrSuccess)
        {
          ShowErrorMessage(err);
        }
      }
      catch (Exception ex)
      {
        return;
      }
    }
    #endregion

    #region Event
    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      string exePath = Path.GetDirectoryName(typeof(MainWindow).Assembly.Location);
      string resPath = exePath + "\\";

      if (LibraryManager.InitLibrary(Path.Combine(resPath, "x64")))
      {
        LibraryManager.Initialize(Path.Combine(resPath, "resource"));
        LibraryManager.SetProgress(Marshal.GetFunctionPointerForDelegate(getPorgress));
        ErrorCode result = LibraryManager.LicenseVerify(Path.Combine(resPath, "license.xml"));
        LibraryManager.SetDocumentAIModel(Path.Combine(resPath, "resource", "models", "documentai.model"), new List<OCRLanguage> { OCRLanguage.e_ENGLISH });

        if (result != ErrorCode.e_ErrSuccess)
        {
          MessageBox.Show("ComPDFKit Conversion SDK Load Failed!");
        }
      }
      else
      {
        MessageBox.Show("ComPDFKit Conversion SDK NativeLibrary Load Failed!");
      }
    }

    private void Input_Click(object sender, RoutedEventArgs e)
    {
      var dlg = new Microsoft.Win32.OpenFileDialog();
      dlg.Filter = "PDF Image Files (*.pdf;*.bmp;*.jpg;*.jpeg;*.png;*.tiff;*.webp)|*.pdf;*.bmp;*.jpg;*.jpeg;*.png;*.tiff;*.webp";
      dlg.Multiselect = true;

      if (dlg.ShowDialog() == true)
      {
        selectedFiles = new List<string>(dlg.FileNames);

        if (selectedFiles.Count > 0)
        {
          InputPath.Text = selectedFiles[0];
        }
        Progress.Text = "";
      }
    }

    private void Output_Click(object sender, RoutedEventArgs e)
    {
      FolderSelectDialog dlg = new FolderSelectDialog();

      if (dlg.ShowDialog())
      {
        OutputPath.Text = dlg.FileName;
      }
    }

    private async void Convert_Click(object sender, RoutedEventArgs e)
    {
      if (selectedFiles == null || selectedFiles.Count == 0)
      {
        MessageBox.Show("Invalid input path!");
        return;
      }

      if (string.IsNullOrEmpty(OutputPath.Text))
      {
        MessageBox.Show("Invalid output path!");
        return;
      }

      Cancel.IsEnabled = true;
      Convert.IsEnabled = false;
      ConvertType.IsEnabled = false;
      ConverterOptions.IsEnabled = false;
      if (Options.EnableOCR)
        LibraryManager.SetOCRLanguage(new List<OCRLanguage> { Options.OCRLanguage });

      foreach (string filePath in selectedFiles)
      {
        InputPath.Text = filePath;
        switch ((ConvertType.SelectedItem as ComboBoxItem).Name)
        {
          case "Word":
            await WordConvert();
            break;

          case "Excel":
            await ExcelConvert();
            break;

          case "Ppt":
            await PptConvert();
            break;

          case "Html":
            await HtmlConvert();
            break;

          case "Rtf":
            await RtfConvert();
            break;

          case "SearchablePDF":
            LibraryManager.SetOCRLanguage(new List<OCRLanguage> { Options.OCRLanguage });
            await PdfConvert();
            break;

          case "Txt":
            await TxtConvert();
            break;

          case "Json":
            await JsonConvert();
            break;

          case "Image":
            await ImageConvert();
            break;

          case "Markdown":
            await MarkdownConvert();
            break;

          default:
            break;
        }
      }

      if (err == ErrorCode.e_ErrSuccess)
        Process.Start(OutputPath.Text);

      Cancel.IsEnabled = false;
      Convert.IsEnabled = true;
      ConvertType.IsEnabled = true;
      ConverterOptions.IsEnabled = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
      CPDFConversion.Cancel();
    }

    private void ConverterOptions_Click(object sender, RoutedEventArgs e)
    {
      ConverterOptionsWindow optionsWindow = new ConverterOptionsWindow(this, (ConvertType.SelectedItem as ComboBoxItem).Name);
      optionsWindow.ShowDialog();
    }
    #endregion
  }
}
