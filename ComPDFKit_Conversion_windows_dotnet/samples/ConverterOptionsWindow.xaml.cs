using ComPDF_Conversion.Common;
using System.Windows;
using System.Windows.Forms;
using System.Xml;

namespace ComPDF_Conversion_Demo
{
    public partial class ConverterOptionsWindow : Window
    {
        private MainWindow parent;

        public ConverterOptionsWindow(MainWindow parent, string type)
        {
            InitializeComponent();
            this.parent = parent;

            ContainAnnotations.Click += (sender, args) =>
            {
                if (ContainAnnotations.IsChecked != null)
                    this.parent.Options.ContainAnnotations = ContainAnnotations.IsChecked.Value;
            };

            CsvFormat.Click += (sender, args) =>
            {
                if (CsvFormat.IsChecked != null) this.parent.Options.CsvFormat = CsvFormat.IsChecked.Value;
            };

            AllContent.Click += (sender, args) =>
            {
                if (AllContent.IsChecked != null) this.parent.Options.AllContent = AllContent.IsChecked.Value;
            };

            ContainImages.Click += (sender, args) =>
            {
                if (ContainImages.IsChecked != null) this.parent.Options.ContainImages = ContainImages.IsChecked.Value;
            };

            ContainPageBackgroundImage.Click += (sender, args) =>
            {
                if (ContainPageBackgroundImage.IsChecked != null) this.parent.Options.ContainPageBackgroundImage = ContainPageBackgroundImage.IsChecked.Value;
            };

            AutoCreateFolder.Click += (sender, args) =>
            {
                if (AutoCreateFolder.IsChecked != null) this.parent.Options.AutoCreateFolder = AutoCreateFolder.IsChecked.Value;
            };

            OutputDocumentPerPage.Click += (sender, args) =>
            {
                if (OutputDocumentPerPage.IsChecked != null) this.parent.Options.OutputDocumentPerPage = OutputDocumentPerPage.IsChecked.Value;
            };

            EnableAiLayout.Click += (sender, args) =>
            {
                if (EnableAiLayout.IsChecked != null) this.parent.Options.EnableAiLayout = EnableAiLayout.IsChecked.Value;
            };

            EnableOCR.Click += (sender, args) =>
            {
                if (EnableOCR.IsChecked != null) this.parent.Options.EnableOCR = EnableOCR.IsChecked.Value;
                OCRLanguage.IsEnabled = (bool)EnableOCR.IsChecked;
                OCROptions.IsEditable = (bool)EnableOCR.IsChecked;
            };

            TxtTableFormat.Click += (sender, args) =>
            {
                if (TxtTableFormat.IsChecked != null) this.parent.Options.TxtTableFormat = TxtTableFormat.IsChecked.Value;
            };

            FormulaToImage.Click += (sender, args) =>
            {
                if (FormulaToImage.IsChecked != null) this.parent.Options.FormulaToImage = FormulaToImage.IsChecked.Value;
            };

            ImagePathEnhance.Click += (sender, args) =>
            {
                if (ImagePathEnhance.IsChecked != null) this.parent.Options.ImagePathEnhance = ImagePathEnhance.IsChecked.Value;
            };

            ContainTables.Click += (sender, args) =>
            {
                if (ContainTables.IsChecked != null) this.parent.Options.ContainTables = ContainTables.IsChecked.Value;
            };

            ImageFormat.SelectionChanged += ImageFormatComboBox_SelectionChanged;
            ImageMode.SelectionChanged += ImageModeComboBox_SelectionChanged;

            ExcelWorksheetOptions.SelectionChanged += ExcelWorksheetOptions_SelectionChanged;
            HtmlOptions.SelectionChanged += HtmlOptions_SelectionChanged;

            ImageRatio.LostFocus += ImageDpi_LostFocus;

            OCRLanguage.SelectionChanged += OCRLanguage_SelectionChanged;
            OCROptions.SelectionChanged += OCROptions_SelectionChanged;

            PageLayoutOptions.SelectionChanged += PageLayoutOptions_SelectionChanged;

            EnableOCR.Checked += EnableOCR_Checked;
            EnableOCR.Unchecked += EnableOCR_Unchecked;

            switch (type)
            {
                case "Word":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainPageBackgroundImagePanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    IsFormulaToImagePanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    PageLayoutOptionsPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 600;
                    break;
                case "Html":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    IsFormulaToImagePanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    PageLayoutOptionsPanel.Visibility = Visibility.Visible;
                    HtmlOptionsPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 600;
                    break;
                case "SearchablePDF":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainPageBackgroundImagePanel.Visibility = Visibility.Visible;
                    IsFormulaToImagePanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    OCRLanguage.IsEnabled = true;
                    ContainPageBackgroundImage.IsEnabled = true;
                    break;
                case "Ppt":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainPageBackgroundImagePanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    IsFormulaToImagePanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 400;
                    break;
                case "Rtf":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    IsFormulaToImagePanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 400;
                    break;
                case "Markdown":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 400;
                    break;
                case "Excel":
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsAutoCreateFolderPanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    IsFormulaToImagePanel.Visibility = Visibility.Visible;
                    IsCsvFormatPanel.Visibility = Visibility.Visible;
                    IsAllContentPanel.Visibility = Visibility.Visible;
                    ExcelWorksheetOptionsPanel.Visibility = Visibility.Visible;
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 600;
                    break;
                case "Txt":
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    IsTxtTableFormatPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 400;
                    break;
                case "Image":
                    ImageRatioPanel.Visibility = Visibility.Visible;
                    ImageFormatPanel.Visibility = Visibility.Visible;
                    ImageModePanel.Visibility = Visibility.Visible;
                    IsImagePathEnhancePanel.Visibility = Visibility.Visible;
                    Height = 400;
                    break;
                case "Json":
                    OCRLanguagePanel.Visibility = Visibility.Visible;
                    OCROptionsPanel.Visibility = Visibility.Visible;
                    IsEnableOCRPanel.Visibility = Visibility.Visible;
                    IsContainImagesPanel.Visibility = Visibility.Visible;
                    IsContainAnnotationsPanel.Visibility = Visibility.Visible;
                    IsEnableAiLayoutPanel.Visibility = Visibility.Visible;
                    IsContainTablesPanel.Visibility = Visibility.Visible;
                    IsOutputDocumentPerPagePanel.Visibility = Visibility.Visible;
                    Height = 400;
                    break;
            }
        }

        private void OCROptions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.OcrOption = (OCROption)OCROptions.SelectedIndex;
        }

        private void HtmlOptions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.htmlOption = (HtmlPageOption)HtmlOptions.SelectedIndex;
        }

        private void ExcelWorksheetOptions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.WorksheetOption = (ExcelWorksheetOption)ExcelWorksheetOptions.SelectedIndex;
        }

        private void PageLayoutOptions_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.LayoutMode = (PageLayoutMode)PageLayoutOptions.SelectedIndex;
        }

        private void CsvFormat_Click(object sender, RoutedEventArgs e)
        {
            throw new System.NotImplementedException();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            OCRLanguage.SelectedIndex = (int)parent.Options.OCRLanguage - 1;
            OCROptions.SelectedIndex = (int)parent.Options.OcrOption;
            ContainAnnotations.IsChecked = parent.Options.ContainAnnotations;
            ContainImages.IsChecked = parent.Options.ContainImages;
            ContainPageBackgroundImage.IsChecked = parent.Options.ContainPageBackgroundImage;
            AutoCreateFolder.IsChecked = parent.Options.AutoCreateFolder;
            OutputDocumentPerPage.IsChecked = parent.Options.OutputDocumentPerPage;
            EnableAiLayout.IsChecked = parent.Options.EnableAiLayout;
            EnableOCR.IsChecked = parent.Options.EnableOCR;
            CsvFormat.IsChecked = parent.Options.CsvFormat;
            AllContent.IsChecked = parent.Options.AllContent;
            TxtTableFormat.IsChecked = parent.Options.TxtTableFormat;
            ImagePathEnhance.IsChecked = parent.Options.ImagePathEnhance;
            ContainTables.IsChecked = parent.Options.ContainTables;
            FormulaToImage.IsChecked = parent.Options.FormulaToImage;

            ImageFormat.SelectedIndex = (int)parent.Options.ImageFormat;
            ImageMode.SelectedIndex = (int)parent.Options.ImageMode;
            ExcelWorksheetOptions.SelectedIndex = (int)parent.Options.WorksheetOption;
            HtmlOptions.SelectedIndex = (int)parent.Options.htmlOption;

            ImageRatio.Text = parent.Options.ImageRatio.ToString();

            PageLayoutOptions.SelectedIndex = (int)parent.Options.LayoutMode;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void ImageFormatComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.ImageFormat = (ImageType)ImageFormat.SelectedIndex;
        }

        private void ImageModeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.ImageMode = (ImageColorMode)ImageMode.SelectedIndex;
        }

        private void ImageDpi_LostFocus(object sender, RoutedEventArgs e)
        {
            float dpi;
            if (float.TryParse(ImageRatio.Text, out dpi) && dpi > 0)
            {
                if (dpi > 100)
                {
                    ImageRatio.Text = "100.0";
                    parent.Options.ImageRatio = 100.0f;
                }
                parent.Options.ImageRatio = dpi;
            }
            else
            {
                ImageRatio.Text = "1.0";
                parent.Options.ImageRatio = 1.0f;
            }
        }

        private void OCRLanguage_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            parent.Options.OCRLanguage = (OCRLanguage)OCRLanguage.SelectedIndex + 1;
        }

        private void EnableOCR_Checked(object sender, RoutedEventArgs e)
        {
            parent.Options.EnableOCR = (bool)EnableOCR.IsChecked;
            OCRLanguage.IsEnabled = (bool)EnableOCR.IsChecked;
            OCROptions.IsEnabled = (bool)EnableOCR.IsChecked;
            EnableAiLayout.IsEnabled = !parent.Options.EnableOCR;
            ContainPageBackgroundImage.IsEnabled = (bool)EnableOCR.IsChecked;
            ContainPageBackgroundImage.IsChecked = (bool)EnableOCR.IsChecked;
            parent.Options.ContainPageBackgroundImage = (bool)EnableOCR.IsChecked;
        }

        private void EnableOCR_Unchecked(object sender, RoutedEventArgs e)
        {
            parent.Options.EnableOCR = (bool)EnableOCR.IsChecked;
            OCRLanguage.IsEnabled = (bool)EnableOCR.IsChecked;
            OCROptions.IsEnabled = (bool)EnableOCR.IsChecked;
            EnableAiLayout.IsEnabled = !parent.Options.EnableOCR;
            ContainPageBackgroundImage.IsEnabled = (bool)EnableOCR.IsChecked;
            ContainPageBackgroundImage.IsChecked = (bool)EnableOCR.IsChecked;
            parent.Options.ContainPageBackgroundImage = (bool)EnableOCR.IsChecked;
        }
    }
}
