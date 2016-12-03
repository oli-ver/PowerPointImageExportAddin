using System;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointExportCustomResolutionAddIn
{
    partial class ImageExportRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ImageExportRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(exportImage);
        }

        private void exportImage(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Exporting current slide as image started");
            var targetSlide = Globals.ImageExportAddin.Application.ActiveWindow.View.Slide;
            var currentPresentation = Globals.ImageExportAddin.Application.ActivePresentation;
            int height = (int)currentPresentation.PageSetup.SlideHeight;
            int width = (int)currentPresentation.PageSetup.SlideWidth;
            int dpiValue = Int32.Parse(resolutionDropDown.SelectedItem.Label);
            System.Diagnostics.Debug.WriteLine("Current dimensions: Height: " + height + " Width: " + width + " DPI: " + dpiValue);
            float newHeight = (height / (float)72) * dpiValue;
            float newWidth =(width / (float)72) * dpiValue;
            System.Diagnostics.Debug.WriteLine("Calculated new dimensions: Height: " + newHeight + " Width: " + newWidth);

            // Let the user choose the path.
            string path = null;
            System.Windows.Forms.SaveFileDialog file = new System.Windows.Forms.SaveFileDialog();
            file.Filter = getFileFilter(imageFormatDropDown.SelectedItem);
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = file.FileName;
            }
            if (path != null)
            {
                System.Diagnostics.Debug.WriteLine("Starting image export to path " + path + " with " + dpiValue + " dpi");
                targetSlide.Export(path, "." + imageFormatDropDown.SelectedItem.Label, (int)newWidth, (int)newHeight);
            }else
            {
                System.Diagnostics.Debug.WriteLine("Import aborted. User did not choose a path to save the file");
            }
        }

        /// <summary>
        /// Generates a file filter for the selected image format
        /// </summary>
        /// <param name="selectedItem"></param>
        /// <returns> file filter (i. e. Image files (*.jpg)|*.jpg)</returns>
        private string getFileFilter(RibbonDropDownItem selectedItem)
        {
            string fileFilter = null;
            switch (selectedItem.Label)
            {
                case "jpg": fileFilter = "Jpg images (*.jpg)|*.jpg";
                    break;
                case "gif": fileFilter = "Gif images (*.gif)|*.gif";
                    break;
                case "png": fileFilter = "PNG images (*.png)|*.png";
                    break;
                case "tif": fileFilter = "TIFF images (*.tif)|*.tif";
                    break;
                case "bmp": fileFilter = "BMP images (*.bmp)|*.bmp";
                    break;
                case "wmf": fileFilter = "Windows Metafile (*.wmf)|*.wmf";
                    break;
                case "emf": fileFilter = "Enhanced Windows Metafile (*.emf)|*.emf";
                    break;
            }
            return fileFilter;
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für Designerunterstützung -
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.imageExportTab = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.resolutionDropDown = this.Factory.CreateRibbonDropDown();
            this.imageFormatDropDown = this.Factory.CreateRibbonDropDown();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.imageExportTab.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // imageExportTab
            // 
            this.imageExportTab.Groups.Add(this.group2);
            this.imageExportTab.Label = "Image Export";
            this.imageExportTab.Name = "imageExportTab";
            // 
            // group2
            // 
            this.group2.Items.Add(this.resolutionDropDown);
            this.group2.Items.Add(this.imageFormatDropDown);
            this.group2.Items.Add(this.button1);
            this.group2.Label = "Define Resolution";
            this.group2.Name = "group2";
            // 
            // resolutionDropDown
            // 
            ribbonDropDownItemImpl1.Label = "50";
            ribbonDropDownItemImpl2.Label = "96";
            ribbonDropDownItemImpl3.Label = "100";
            ribbonDropDownItemImpl4.Label = "150";
            ribbonDropDownItemImpl5.Label = "200";
            ribbonDropDownItemImpl6.Label = "250";
            ribbonDropDownItemImpl7.Label = "300";
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl1);
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl2);
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl3);
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl4);
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl5);
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl6);
            this.resolutionDropDown.Items.Add(ribbonDropDownItemImpl7);
            this.resolutionDropDown.Label = "Resolution";
            this.resolutionDropDown.Name = "resolutionDropDown";
            // 
            // imageFormatDropDown
            // 
            ribbonDropDownItemImpl8.Label = "jpg";
            ribbonDropDownItemImpl9.Label = "gif";
            ribbonDropDownItemImpl10.Label = "png";
            ribbonDropDownItemImpl11.Label = "tif";
            ribbonDropDownItemImpl12.Label = "bmp";
            ribbonDropDownItemImpl13.Label = "wmf";
            ribbonDropDownItemImpl14.Label = "emf";
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl8);
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl9);
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl10);
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl11);
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl12);
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl13);
            this.imageFormatDropDown.Items.Add(ribbonDropDownItemImpl14);
            this.imageFormatDropDown.Label = "Image Format";
            this.imageFormatDropDown.Name = "imageFormatDropDown";
            // 
            // button1
            // 
            this.button1.Label = "Export";
            this.button1.Name = "button1";
            // 
            // ImageExportRibbon
            // 
            this.Name = "ImageExportRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.imageExportTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.imageExportTab.ResumeLayout(false);
            this.imageExportTab.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab imageExportTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown resolutionDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown imageFormatDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal ImageExportRibbon Ribbon1
        {
            get { return this.GetRibbon<ImageExportRibbon>(); }
        }
    }
}
