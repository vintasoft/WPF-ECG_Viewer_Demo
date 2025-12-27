using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

using Microsoft.Win32;

using Vintasoft.Imaging;
using Vintasoft.Imaging.Codecs.Decoders;
using Vintasoft.Imaging.Dicom.Wpf.UI.VisualTools;
using Vintasoft.Imaging.Metadata;
using Vintasoft.Imaging.Wpf.Print;

using WpfDemosCommonCode;
using WpfDemosCommonCode.Imaging.Codecs;

namespace WpfEcgViewerDemo
{
    /// <summary>
    /// Main window of ECG viewer demo.
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Fields

        /// <summary>
        /// Template of the application title.
        /// </summary>
        string _titlePrefix = "VintaSoft WPF ECG Viewer Demo v" + ImagingGlobalSettings.ProductVersion + " - {0}";

        /// <summary>
        /// The print manager.
        /// </summary>
        WpfImagePrintManager _printManager;


        #region File Dialogs

        /// <summary>
        /// The open file dialog for DICOM file.
        /// </summary>
        OpenFileDialog _openDicomFileDialog = new OpenFileDialog();

        /// <summary>
        /// The save file dialog for DICOM file.
        /// </summary>
        SaveFileDialog _saveDicomFileDialog = new SaveFileDialog();

        #endregion


        #region Hot keys

        public static RoutedCommand _openCommand = new RoutedCommand();
        public static RoutedCommand _saveCommand = new RoutedCommand();
        public static RoutedCommand _closeCommand = new RoutedCommand();
        public static RoutedCommand _printCommand = new RoutedCommand();
        public static RoutedCommand _exitCommand = new RoutedCommand();
        public static RoutedCommand _aboutCommand = new RoutedCommand();

        #endregion

        #endregion



        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            imageViewerToolStrip1.ImageViewer = imageViewer1;

            this.Title = string.Format(_titlePrefix, "(Untitled)");

            // set the initial directory in open dicom file dialog
            _openDicomFileDialog.Filter = "DICOM files|*.dcm;*.dic;*.acr|All files|*.*";
            DemosTools.SetTestFilesFolder(_openDicomFileDialog);

            imageViewer1.ImageRenderingSettings = new EcgRenderingSettings();
            imageViewer1.ImageRenderingSettings.Changed += ImageRenderingSettings_Changed;
            imageViewer1.FocusedIndexChanged += ImageViewer1_FocusedIndexChanged;

            // create visual tool
            WpfEcgVisualTool visualTool = new WpfEcgVisualTool();
            visualTool.SelectionChanged += VisualTool_SelectionChanged;
            visualTool.IsEnabled = false;
            imageViewer1.VisualTool = visualTool;

            // create the print manager
            _printManager = new WpfImagePrintManager();
            _printManager.Images = imageViewer1.Images;
            _printManager.PrintScaleMode = Vintasoft.Imaging.Print.PrintScaleMode.BestFit;
            _printManager.PagePadding = new Thickness(10);

            UpdateUI();
        }

        #endregion



        #region Properties

        bool _isFileOpening = false;
        /// <summary>
        /// Gets or sets a value indicating whether the file is opening.
        /// </summary>
        private bool IsFileOpening
        {
            get
            {
                return _isFileOpening;
            }
            set
            {
                _isFileOpening = value;

                InvokeUpdateUI();
            }
        }

        /// <summary>
        /// Gets the <see cref="EcgRenderingSettings"/>.
        /// </summary>
        private EcgRenderingSettings EcgRenderingSettings
        {
            get
            {
                return (EcgRenderingSettings)imageViewer1.ImageRenderingSettings;
            }
        }

        #endregion



        #region Methods

        #region PRIVATE

        #region Window

        /// <summary>
        /// Handles the Closed event of Window object.
        /// </summary>
        private void Window_Closed(object sender, EventArgs e)
        {
            _printManager.Dispose();
        }

        #endregion


        #region UI state

        /// <summary>
        /// Updates the user interface safely.
        /// </summary>
        private void InvokeUpdateUI()
        {
            if (Dispatcher.Thread == Thread.CurrentThread)
                UpdateUI();
            else
                Dispatcher.Invoke(new UpdateUIDelegate(UpdateUI));
        }

        /// <summary>
        /// Updates the user interface of this form.
        /// </summary>
        private void UpdateUI()
        {
            bool isOpening = IsFileOpening;
            bool isOpened = imageViewer1.Image != null;

            openMenuItem.IsEnabled = !isOpening;
            saveMenuItem.IsEnabled = !isOpening && isOpened;
            printMenuItem.IsEnabled = !isOpening && isOpened;
            closeMenuItem.IsEnabled = !isOpening && isOpened;
            imageViewerToolStrip1.SaveButtonEnabled = saveMenuItem.IsEnabled;
            imageViewerToolStrip1.PrintButtonEnabled = printMenuItem.IsEnabled;
            imageViewerToolStrip1.IsNavigationEnabled = !isOpening && isOpened;


            gainMenuItem.IsEnabled = !isOpening && isOpened;
            foreach (MenuItem item in gainMenuItem.Items)
                item.IsChecked = false;
            switch (EcgRenderingSettings.MillimetersPerMillivolt)
            {
                case 5:
                    gain5mmMenuItem.IsChecked = true;
                    break;

                case 10:
                    gain10mmMenuItem.IsChecked = true;
                    break;

                case 20:
                    gain20mmMenuItem.IsChecked = true;
                    break;

                case 40:
                    gain40mmMenuItem.IsChecked = true;
                    break;
            }


            gridTypeMenuItem.IsEnabled = !isOpening && isOpened;
            foreach (MenuItem item in gridTypeMenuItem.Items)
                item.IsChecked = false;
            if (EcgRenderingSettings.MajorGridThickness == 0 && EcgRenderingSettings.MinorGridThickness == 0)
            {
                gridTypeNoneMenuItem.IsChecked = true;
            }
            else if (EcgRenderingSettings.MajorGridThickness == 1)
            {
                if (EcgRenderingSettings.MinorGridThickness == 0)
                    gridType5mmMenuItem.IsChecked = true;
                else
                    gridType1mmMenuItem.IsChecked = true;
            }


            colorMenuItem.IsEnabled = !isOpening && isOpened;


            caliperMenuItem.IsEnabled = !isOpening && isOpened;
            foreach (MenuItem item in caliperMenuItem.Items)
                item.IsChecked = false;
            WpfEcgVisualTool visualTool = (WpfEcgVisualTool)imageViewer1.VisualTool;
            if (visualTool.IsEnabled)
            {
                if (visualTool.InterationMode == EcgInterationMode.Duration)
                    caliperDurationMenuItem.IsChecked = true;
                else
                    caliperDurationAndMVMenuItem.IsChecked = true;
            }
            else
            {
                caliperOffMenuItem.IsChecked = true;
            }
        }

        #endregion


        #region 'File' menu

        /// <summary>
        /// Handles the Click event of openMenuItem object.
        /// </summary>
        private void openMenuItem_Click(object sender, RoutedEventArgs e)
        {
            OpenDicomFile();
        }

        /// <summary>
        /// Handles the Click event of saveMenuItem object.
        /// </summary>
        private void saveMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SaveDicomFile();
        }

        /// <summary>
        /// Handles the Click event of closeMenuItem object.
        /// </summary>
        private void closeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            CloseDicomFile();
        }

        /// <summary>
        /// Handles the Click event of printMenuItem object.
        /// </summary>
        private void printMenuItem_Click(object sender, RoutedEventArgs e)
        {
            PrintDicomFile();
        }

        /// <summary>
        /// Handles the Click event of exitMenuItem object.
        /// </summary>
        private void exitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        #endregion


        #region 'View' menu

        /// <summary>
        /// Handles the Click event of gainMenuItem object.
        /// </summary>
        private void gainMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EcgRenderingSettings.BeginInit();
            if (sender == gain5mmMenuItem)
                EcgRenderingSettings.MillimetersPerMillivolt = 5;
            else if (sender == gain10mmMenuItem)
                EcgRenderingSettings.MillimetersPerMillivolt = 10;
            else if (sender == gain20mmMenuItem)
                EcgRenderingSettings.MillimetersPerMillivolt = 20;
            else if (sender == gain40mmMenuItem)
                EcgRenderingSettings.MillimetersPerMillivolt = 40;
            EcgRenderingSettings.EndInit();
        }

        /// <summary>
        /// Handles the Click event of gridTypeMenuItem object.
        /// </summary>
        private void gridTypeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            EcgRenderingSettings.BeginInit();
            if (sender == gridTypeNoneMenuItem)
            {
                EcgRenderingSettings.MinorGridThickness = 0;
                EcgRenderingSettings.MajorGridThickness = 0;
            }
            else if (sender == gridType1mmMenuItem)
            {
                EcgRenderingSettings.MinorGridThickness = 0.5f;
                EcgRenderingSettings.MajorGridThickness = 1;
            }
            else if (sender == gridType5mmMenuItem)
            {
                EcgRenderingSettings.MinorGridThickness = 0;
                EcgRenderingSettings.MajorGridThickness = 1;
            }
            EcgRenderingSettings.EndInit();
        }

        /// <summary>
        /// Handles the Click event of colorMenuItem object.
        /// </summary>
        private void colorMenuItem_Click(object sender, RoutedEventArgs e)
        {
            colorRedBlackMenuItem.IsChecked = false;
            colorBlueBlackMenuItem.IsChecked = false;
            colorGreenBlackMenuItem.IsChecked = false;
            colorGrayGreenMenuItem.IsChecked = false;
            ((MenuItem)sender).IsChecked = true;

            EcgRenderingSettings.BeginInit();
            WpfEcgVisualTool visualTool = (WpfEcgVisualTool)imageViewer1.VisualTool;
            if (sender == colorRedBlackMenuItem)
            {
                imageViewer1.Background = Brushes.White;
                EcgRenderingSettings.BackgroundColor = System.Drawing.Color.White;
                EcgRenderingSettings.MajorGridColor = System.Drawing.Color.FromArgb(255, 187, 187);
                EcgRenderingSettings.MinorGridColor = System.Drawing.Color.FromArgb(255, 229, 229);
                EcgRenderingSettings.SignalColor = System.Drawing.Color.Black;
                EcgRenderingSettings.LegendFontColor = System.Drawing.Color.Black;
                visualTool.SelectionPen = new Pen(Brushes.Black, 1);
            }
            else if (sender == colorBlueBlackMenuItem)
            {
                imageViewer1.Background = Brushes.White;
                EcgRenderingSettings.BackgroundColor = System.Drawing.Color.White;
                EcgRenderingSettings.MajorGridColor = System.Drawing.Color.FromArgb(187, 187, 255);
                EcgRenderingSettings.MinorGridColor = System.Drawing.Color.FromArgb(229, 229, 255);
                EcgRenderingSettings.SignalColor = System.Drawing.Color.Black;
                EcgRenderingSettings.LegendFontColor = System.Drawing.Color.Black;
                visualTool.SelectionPen = new Pen(Brushes.Black, 1);
            }
            else if (sender == colorGreenBlackMenuItem)
            {
                imageViewer1.Background = Brushes.White;
                EcgRenderingSettings.BackgroundColor = System.Drawing.Color.White;
                EcgRenderingSettings.MajorGridColor = System.Drawing.Color.FromArgb(28, 255, 28);
                EcgRenderingSettings.MinorGridColor = System.Drawing.Color.FromArgb(204, 255, 204);
                EcgRenderingSettings.SignalColor = System.Drawing.Color.Black;
                EcgRenderingSettings.LegendFontColor = System.Drawing.Color.Black;
                visualTool.SelectionPen = new Pen(Brushes.Black, 1);
            }
            else if (sender == colorGrayGreenMenuItem)
            {
                imageViewer1.Background = Brushes.Black;
                EcgRenderingSettings.BackgroundColor = System.Drawing.Color.Black;
                EcgRenderingSettings.MajorGridColor = System.Drawing.Color.Gray;
                EcgRenderingSettings.MinorGridColor = System.Drawing.Color.FromArgb(96, 96, 96);
                EcgRenderingSettings.SignalColor = System.Drawing.Color.Lime;
                EcgRenderingSettings.LegendFontColor = System.Drawing.Color.Lime;
                visualTool.SelectionPen = new Pen(Brushes.Lime, 1);
            }
            EcgRenderingSettings.EndInit();
        }

        /// <summary>
        /// Handles the Click event of caliperMenuItem object.
        /// </summary>
        private void caliperMenuItem_Click(object sender, RoutedEventArgs e)
        {
            WpfEcgVisualTool visualTool = (WpfEcgVisualTool)imageViewer1.VisualTool;

            if (sender == caliperOffMenuItem)
            {
                visualTool.IsEnabled = false;
            }
            else if (sender == caliperDurationMenuItem)
            {
                visualTool.IsEnabled = true;
                visualTool.InterationMode = EcgInterationMode.Duration;
            }
            else if (sender == caliperDurationAndMVMenuItem)
            {
                visualTool.IsEnabled = true;
                visualTool.InterationMode = EcgInterationMode.DurationAndVoltage;
            }
            InvokeUpdateUI();
        }

        #endregion


        #region 'Help' menu

        /// <summary>
        /// Handles the Click event of aboutMenuItem object.
        /// </summary>
        private void aboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder description = new StringBuilder();

            description.AppendLine("This project demonstrates how to preview electrocardiogram from DICOM file and allows to:");
            description.AppendLine();
            description.AppendLine("- Open DICOM file with electrocardiogram.");
            description.AppendLine();
            description.AppendLine("- View and print electrocardiogram.");
            description.AppendLine();
            description.AppendLine("- Measure electrocardiogram using caliper.");
            description.AppendLine();

            description.AppendLine();
            description.AppendLine("The project is available in C# and VB.NET for Visual Studio .NET.");

            WpfAboutBoxBaseWindow dlg = new WpfAboutBoxBaseWindow("vsdicom-dotnet");
            dlg.Description = description.ToString();
            dlg.Owner = this;
            dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dlg.ShowDialog();
        }

        #endregion


        #region Hot keys

        /// <summary>
        /// Handles the CanExecute event of openCommandBinding object.
        /// </summary>
        private void openCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = openMenuItem.IsEnabled;
        }

        /// <summary>
        /// Handles the CanExecute event of saveCommandBinding object.
        /// </summary>
        private void saveCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = saveMenuItem.IsEnabled;
        }

        /// <summary>
        /// Handles the CanExecute event of closeCommandBinding object.
        /// </summary>
        private void closeCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = closeMenuItem.IsEnabled;
        }

        /// <summary>
        /// Handles the CanExecute event of printCommandBinding object.
        /// </summary>
        private void printCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = printMenuItem.IsEnabled;
        }

        /// <summary>
        /// Handles the CanExecute event of exitCommandBinding object.
        /// </summary>
        private void exitCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = exitMenuItem.IsEnabled;
        }

        #endregion


        #region Visual Tools

        /// <summary>
        /// Handles the SelectionChanged event of VisualTool object.
        /// </summary>
        private void VisualTool_SelectionChanged(object sender, PropertyChangedEventArgs<Vintasoft.Primitives.VintasoftRect> e)
        {
            if (e.NewValue.Width != 0 || e.NewValue.Height != 0)
            {
                StringBuilder stringBuilder = new StringBuilder();

                stringBuilder.AppendFormat("{0:F0} ms", e.NewValue.Width * 1000);

                if (e.NewValue.Height != 0)
                    stringBuilder.AppendFormat(", {0:F0} μV", e.NewValue.Height * 1000);

                ecgVisualToolSelectionLabel.Text = stringBuilder.ToString();
            }
            else
            {
                ecgVisualToolSelectionLabel.Text = string.Empty;
            }
        }

        #endregion


        #region File manipulation

        /// <summary>
        /// Opens a DICOM file.
        /// </summary>
        private void OpenDicomFile()
        {
            // if file can be opened
            if (_openDicomFileDialog.ShowDialog() == true)
            {
                IsFileOpening = true;
                try
                {
                    // open file
                    VintasoftImage image = new VintasoftImage(_openDicomFileDialog.FileName);

                    // if opened is ECG file
                    if (image.Metadata.MetadataTree.FindChildNode<EcgMetadata>() != null)
                    {
                        // remove images from image viewer
                        imageViewer1.Images.ClearAndDisposeItems();
                        // add image to image viewer
                        imageViewer1.Images.Add(image);

                        // update header of form
                        this.Title = string.Format(_titlePrefix, Path.GetFileName(_openDicomFileDialog.FileName));
                    }
                    else
                    {
                        // remove image
                        image.Dispose();
                        throw new InvalidOperationException(string.Format(
                            "The file '{0}' can't be opened because file does not contain the electrocardiogram data.",
                            _openDicomFileDialog.FileName));
                    }
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
                finally
                {
                    IsFileOpening = false;
                }
            }
        }

        /// <summary>
        /// Saves an electrocardiogram image to an image file.
        /// </summary>
        private void SaveDicomFile()
        {
            CodecsFileFilters.SetFilters(_saveDicomFileDialog, false);
            // if file is selected in "Save file" dialog
            if (_saveDicomFileDialog.ShowDialog() == true)
            {
                try
                {
                    string saveFilename = Path.GetFullPath(_saveDicomFileDialog.FileName);

                    // save image collection to a file
                    imageViewer1.Images.SaveAsync(saveFilename);
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
            }
        }

        /// <summary>
        /// Closes a DICOM file.
        /// </summary>
        private void CloseDicomFile()
        {
            // remove images from image viewer
            imageViewer1.Images.ClearAndDisposeItems();
            UpdateUI();
        }

        /// <summary>
        /// Prints a DICOM file.
        /// </summary>
        private void PrintDicomFile()
        {
            PrintDialog printDialog = _printManager.PrintDialog;
            printDialog.MinPage = 1;
            printDialog.MaxPage = (uint)_printManager.Images.Count;

            // show print dialog and
            // start print if dialog results is OK
            if (printDialog.ShowDialog() == true)
            {
                try
                {
                    _printManager.Print(Title);
                }
                catch (Exception ex)
                {
                    DemosTools.ShowErrorMessage(ex);
                }
            }
        }

        #endregion


        /// <summary>
        /// Handles the Changed event of ImageRenderingSettings object.
        /// </summary>
        private void ImageRenderingSettings_Changed(object sender, EventArgs e)
        {
            InvokeUpdateUI();
        }

        /// <summary>
        /// Handles the FocusedIndexChanged event of ImageViewer1 object.
        /// </summary>
        private void ImageViewer1_FocusedIndexChanged(object sender, PropertyChangedEventArgs<int> e)
        {
            StringBuilder builder = new StringBuilder();

            if (imageViewer1.Image != null)
            {
                PageMetadata pageMetadata = imageViewer1.Image.Metadata.MetadataTree;

                AppendMetadataNodeValue(builder, pageMetadata, "Study Date", "Study Date");
                AppendMetadataNodeValue(builder, pageMetadata, "Study Time", "Study Time");
                builder.AppendLine();
                AppendMetadataNodeValue(builder, pageMetadata, "Patient Name", "Patient's Name");
                AppendMetadataNodeValue(builder, pageMetadata, "Patient ID", "Patient ID");
                AppendMetadataNodeValue(builder, pageMetadata, "Patient Birth Date", "Patient's Birth Date");
                AppendMetadataNodeValue(builder, pageMetadata, "Patient Age", "Patient's Age");
                AppendMetadataNodeValue(builder, pageMetadata, "Patient Gender", "Patient's Sex");
            }

            fileMetadataTextBox.Text = builder.ToString();
        }

        /// <summary>
        /// Handles the OpenFile event of imageViewerToolStrip1 object.
        /// </summary>
        private void imageViewerToolStrip1_OpenFile(object sender, EventArgs e)
        {
            OpenDicomFile();
        }

        /// <summary>
        /// Handles the SaveFile event of imageViewerToolStrip1 object.
        /// </summary>
        private void imageViewerToolStrip1_SaveFile(object sender, EventArgs e)
        {
            SaveDicomFile();
        }

        /// <summary>
        /// Handles the Print event of imageViewerToolStrip1 object.
        /// </summary>
        private void imageViewerToolStrip1_Print(object sender, EventArgs e)
        {
            PrintDicomFile();
        }

        /// <summary>
        /// Appends the metadata node value to the specified string builder.
        /// </summary>
        /// <param name="builder">The string builder.</param>
        /// <param name="pageMetadata">The page metadata.</param>
        /// <param name="nodeDescription">The metadata node description.</param>
        /// <param name="nodeName">The metadata node name.</param>
        private void AppendMetadataNodeValue(
            StringBuilder builder,
            PageMetadata pageMetadata,
            string nodeDescription,
            string nodeName)
        {
            // find the metadata node
            MetadataNode node = pageMetadata.FindChildNode<MetadataNode>(nodeName);

            // if metadata node has value
            if (node != null && node.Value != null)
            {
                // add description
                builder.Append(nodeDescription);
                builder.Append(":");
                if (nodeDescription.Length < 24)
                    builder.Append(new string(' ', 24 - nodeDescription.Length));

                // if metadata node contains date time
                if (node.Value is DateTime)
                    builder.Append(((DateTime)node.Value).ToShortDateString());
                else
                    builder.Append(node.Value.ToString());

                builder.AppendLine();
            }
        }

        #endregion

        #endregion



        #region Delegates

        /// <summary>
        /// Represents the <see cref="UpdateUI"/> method.
        /// </summary>
        delegate void UpdateUIDelegate();

        #endregion

    }
}
