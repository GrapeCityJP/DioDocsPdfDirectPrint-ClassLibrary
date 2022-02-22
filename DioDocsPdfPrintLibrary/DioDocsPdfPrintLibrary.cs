using GrapeCity.Documents.Common;
using GrapeCity.Documents.Drawing;
using GrapeCity.Documents.Imaging.Windows;
using GrapeCity.Documents.Pdf;
using GrapeCity.Documents.Pdf.Renderer;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.Numerics;
using System.Runtime.InteropServices;
using D2D = GrapeCity.Documents.DX.Direct2D;
using D3D = GrapeCity.Documents.DX.Direct3D11;
using DW = GrapeCity.Documents.DX.DirectWrite;
using DX = GrapeCity.Documents.DX;
using DXGI = GrapeCity.Documents.DX.DXGI;
using STG = GrapeCity.Documents.DX.Storage;
using WIC = GrapeCity.Documents.DX.WIC;

namespace DioDocsPdfPrintLibrary
{
    /// <summary>
    /// Specifies how pages are scaled when printed.
    /// </summary>
    public enum PageScaling
    {
        /// <summary>
        /// Pages are enlarged or made smaller if needed to fit paper.
        /// </summary>
        FitToPaper,
        /// <summary>
        /// Pages are enlarged or made smaller if needed to fit printable page bounds.
        /// </summary>
        FitToPrintableArea,
    }

    /// <summary>
    /// Provides properties and methods that enable printing
    /// <see cref="GcPdfDocument"/> objects on Windows systems using Direct2D.
    /// </summary>
    public class GcPdfPrintManager
    {
        #region Data members
        private const PageScaling c_DefPageScaling = PageScaling.FitToPaper;

        private PrinterSettings _printerSettings;
        private PageSettings _pageSettings;
        private OutputRange _outputRange;
        private string _printJobName;
        private PageScaling _pageScaling = PageScaling.FitToPaper;
        private bool _autoRotate = true;
        private RenderingCache _renderingCache;
        private bool? _autoCollate;
        private GcPdfDocument _doc;
        #endregion

        #region Protected
        protected bool OnLongOperation(double complete, bool canCancel)
        {
            if (LongOperation != null)
                return LongOperation.Invoke(complete, canCancel);
            return true;
        }
        #endregion

        #region Public properties
        /// <summary>
        /// Gets or sets a value indicating whether the PrinterSettings.Collate property
        /// should be processed by the <see cref="GcPdfPrintManager"/> or by the printer driver.
        /// By default this property is null, the GcPdfPrintManager will try to determine
        /// does printer driver supports collate or not.
        /// </summary>
        public bool? AutoCollate
        {
            get { return _autoCollate; }
            set { _autoCollate = value; }
        }

        /// <summary>
        /// Gets or sets the <see cref="GcPdfDocument"/> object to print.
        /// </summary>
        public GcPdfDocument Doc
        {
            get { return _doc; }
            set { _doc = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether pages should be auto-rotated to better fit the paper during printing.
        /// <para>The default is <see langword="true"/>.</para>
        /// </summary>
        public bool AutoRotate
        {
            get { return _autoRotate; }
            set { _autoRotate = value; }
        }

        /// <summary>
        /// Gets or sets the <see cref="System.Drawing.Printing.PrinterSettings"/> object defining the print parameters.
        /// If <see langword="null"/>, default printer settings will be used.
        /// </summary>
        public PrinterSettings PrinterSettings
        {
            get { return _printerSettings; }
            set { _printerSettings = value; }
        }

        /// <summary>
        /// Gets or sets the <see cref="System.Drawing.Printing.PageSettings"/> object defining the page settings (page size, orientation and so on).
        /// if <see langword="null"/>, default page settings will be used.
        /// </summary>
        public PageSettings PageSettings
        {
            get { return _pageSettings; }
            set { _pageSettings = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating how pages are scaled during printing.
        /// <para>
        /// The default value is <see cref="PageScaling.FitToPaper"/>.
        /// </para>
        /// </summary>
        [DefaultValue(c_DefPageScaling)]
        public PageScaling PageScaling
        {
            get { return _pageScaling; }
            set { _pageScaling = value; }
        }

        /// <summary>
        /// Gets or sets the <see cref="Common.OutputRange"/> object specifying the range of pages to print.
        /// If <see langword="null"/>, the range specified in <see cref="PrinterSettings"/> will be used.
        /// </summary>
        public OutputRange OutputRange
        {
            get { return _outputRange; }
            set { _outputRange = value; }
        }

        /// <summary>
        /// Gets or sets the name for the print job.
        /// </summary>
        public string PrintJobName
        {
            get { return _printJobName; }
            set { _printJobName = value; }
        }

        /// <summary>
        /// Gets or sets the <see cref="Pdf.RenderingCache"/> object to use while rendering.
        /// This can be <see langword="null"/>.
        /// </summary>
        public RenderingCache RenderingCache
        {
            get { return _renderingCache; }
            set { _renderingCache = value; }
        }
        #endregion

        #region Public static
        /// <summary>
        /// Some printer models are not handled correctly by the printing subsystem.
        /// Trying to use <see cref="GcPdfPrintManager"/> with such printers will fail
        /// due to reasons outside of GcPdfPrintManager's control.
        /// Use this method to check whether a particular printer can be used,
        /// note that if it fails that is not GcPdfPrintManager's fault.
        /// </summary>
        /// <param name="printerName">The printer to check.</param>
        /// <returns><see langword="true"/> if the printer is OK, <see langword="false"/> otherwise.</returns>
        public static bool TestPrinter(string printerName)
        {
            try
            {
                var ps = new PrinterSettings() { PrinterName = printerName };
                IntPtr hDevMode = ps.GetHdevmode();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Swaps two values.
        /// </summary>
        /// <typeparam name="T">The type of values.</typeparam>
        public static void Swap<T>(ref T x, ref T y)
        {
            T temp = x;
            x = y;
            y = temp;
        }

        /// <summary>
        /// Returns the paper rotation angle.
        /// </summary>
        /// <param name="pageRotationAngle">The page rotation angle, in degrees.</param>
        /// <returns>The paper rotation angle, in degrees.</returns>
        public static int PaperRotationAngle(int pageRotationAngle)
        {
            if (pageRotationAngle == 90)
                return 270;
            else if (pageRotationAngle == 270)
                return 90;
            else if (pageRotationAngle == 0)
                return 0;
            else
                throw new ArgumentOutOfRangeException("Valid values are 90 & 270.");
        }

        /// <summary>
        /// Tests whether a page should be rotated to better fit paper.
        /// </summary>
        /// <param name="paperSize">The paper size.</param>
        /// <param name="pageSize">The page size.</param>
        /// <returns><b>true</b> if the page should be rotated, <b>false</b> otherwise.</returns>
        public static bool ShouldRotate(SizeF paperSize, SizeF pageSize)
        {
            return (paperSize.Width > paperSize.Height) != (pageSize.Width > pageSize.Height);
        }

        /// <summary>
        /// Swaps the width and height of a <see cref="Size"/> structure (rotates 90 degrees).
        /// </summary>
        /// <param name="s">The <see cref="Size"/> to rotate.</param>
        /// <returns>The newly created <see cref="Size"/> with width and height swapped.</returns>
        public static SizeF RotateSize(SizeF s)
        {
            float t = s.Width;
            s.Width = s.Height;
            s.Height = t;
            return s;
        }

        /// <summary>
        /// Rotates a paper size and the printable area within it by the specified angle.
        /// </summary>
        /// <param name="angle">The rotation angle, counterclockwise (valid values are <b>90</b> and <b>270</b>).</param>
        /// <param name="paperSize">The paper size.</param>
        /// <param name="printableArea">The printable area.</param>
        public static void RotatePaper(int angle,
            ref SizeF paperSize, ref RectangleF printableArea)
        {
            if (angle == 90)
            {
                printableArea = new RectangleF(
                    printableArea.Top,
                    paperSize.Width - printableArea.Right,
                    printableArea.Height, printableArea.Width);
                paperSize = RotateSize(paperSize);
            }
            else if (angle == 270)
            {
                printableArea = new RectangleF(
                    paperSize.Height - printableArea.Bottom,
                    printableArea.Left,
                    printableArea.Height, printableArea.Width);
                paperSize = RotateSize(paperSize);
            }
            else if (angle != 0)
                throw new ArgumentOutOfRangeException("Valid values are 90 & 270.");
        }
        #endregion

        #region Public
        /// <summary>
        /// Prints the document.
        /// </summary>
        public void Print()
        {
            PrinterSettings printerSettings = _printerSettings;
            if (printerSettings == null)
                printerSettings = new PrinterSettings();
            bool autoCollate;
            if (_autoCollate.HasValue)
                autoCollate = _autoCollate.Value;
            else
            {
                // check does printer support DM_COLLATE and DM_COPIES properties or not
                IntPtr hdm = printerSettings.GetHdevmode();
                IntPtr lhdm = GlobalLock(hdm);
                DEVMODE dm = (DEVMODE)Marshal.PtrToStructure(lhdm, typeof(DEVMODE));
                if ((dm.dmFields & DM.DM_COLLATE) != 0 && (dm.dmFields & DM.DM_COPIES) != 0)
                    autoCollate = false;
                else
                    autoCollate = true;
                GlobalUnlock(hdm);
                GlobalFree(hdm);
            }

            int printCopies = printerSettings.Copies;
            if (autoCollate)
            {
                printCopies = printerSettings.Copies;
                printerSettings.Copies = 1;
            }
            else
            {
                printCopies = 1;
            }
            PageSettings pageSettings;
            IntPtr hDevMode = printerSettings.GetHdevmode();
            if (_pageSettings != null)
            {
                pageSettings = _pageSettings;
                pageSettings.CopyToHdevmode(hDevMode);
                printerSettings.SetHdevmode(hDevMode);
            }
            else
            {
                pageSettings = printerSettings.DefaultPageSettings;
            }
            // * 0.96 - used to convert to DIPs
            RectangleF printableArea = new RectangleF(
                pageSettings.PrintableArea.X * 0.96f,
                pageSettings.PrintableArea.Y * 0.96f,
                pageSettings.PrintableArea.Width * 0.96f,
                pageSettings.PrintableArea.Height * 0.96f);
            PaperSize paperSize = pageSettings.PaperSize;
            float printerPageWidthPx = paperSize.Width * 0.96f;
            float printerPageHeightPx = paperSize.Height * 0.96f;
            if (pageSettings.Landscape)
            {
                Swap(ref printerPageWidthPx, ref printerPageHeightPx);
                printableArea = new RectangleF(printableArea.Y, printableArea.X, printableArea.Height, printableArea.Width);
            }
            //
            OutputRange outputRange = _outputRange;
            if (outputRange == null)
            {
                int pMin, pMax;
                switch (printerSettings.PrintRange)
                {
                    case PrintRange.SomePages:
                        pMin = Math.Min(Math.Max(printerSettings.FromPage, 1), _doc.Pages.Count);
                        pMax = Math.Min(Math.Max(printerSettings.ToPage, 1), _doc.Pages.Count);
                        break;
                    default:
                        pMin = 1;
                        pMax = _doc.Pages.Count;
                        break;
                }
                outputRange = new OutputRange(pMin, pMax);
            }
            //
            IntPtr lockedDevMode = GlobalLock(hDevMode);
            DEVMODE devMode = (DEVMODE)Marshal.PtrToStructure(lockedDevMode, typeof(DEVMODE));
            int dpi = 150;
            if ((devMode.dmFields & DM.DM_PRINTQUALITY) != 0 && devMode.dmPrintQuality > 0)
            {
                dpi = devMode.dmPrintQuality;
            }
            else if ((devMode.dmFields & DM.DM_YRESOLUTION) != 0)
            {
                dpi = devMode.dmYResolution;
            }

            STG.ComStream jobPrintTicketStream = null;
            try
            {
                jobPrintTicketStream = CreatePrintTicketFromDevMode(printerSettings.PrinterName, lockedDevMode, devMode.dmSize + devMode.dmDriverExtra);

                Print(printerSettings.PrinterName,
                    jobPrintTicketStream,
                    printerPageWidthPx,
                    printerPageHeightPx,
                    printableArea,
                    dpi,
                    autoCollate,
                    printCopies,
                    printerSettings.Collate,
                    outputRange);
            }
            finally
            {
                if (jobPrintTicketStream != null)
                    jobPrintTicketStream.Dispose();
                GlobalUnlock(hDevMode);
                GlobalFree(hDevMode);
            }
        }
        #endregion

        #region Private
        private STG.ComStream CreatePrintTicketFromDevMode(string printerName, IntPtr lockedDevMode, int devModeSize)
        {
            IntPtr istream = IntPtr.Zero;
            DX.HResult hr = CreateStreamOnHGlobal(IntPtr.Zero, true, ref istream);
            hr.CheckError();
            if (istream == IntPtr.Zero)
            {
                // From:
                // https://msdn.microsoft.com/ru-ru/library/windows/desktop/aa378980(v=vs.85).aspx
                // it looks like istream cannot be null, this just for safety.
                throw new InvalidOperationException();
            }

            STG.ComStream result = new STG.ComStream(istream);
            IntPtr hProvider = IntPtr.Zero;
            try
            {
                hr = PTOpenProvider(printerName, 1, ref hProvider);
                hr.CheckError();

                hr = PTConvertDevModeToPrintTicket(hProvider, devModeSize, lockedDevMode, 2 /* kPTJobScope */, istream);
                hr.CheckError();

                return result;
            }
            catch
            {
                result.Dispose();
                throw;
            }
            finally
            {
                if (hProvider != IntPtr.Zero)
                    PTCloseProvider(hProvider);
            }
        }

        private void AlignInRect(RectangleF rect, float width, float height,
            out float offsX, out float offsY, out float scaleX, out float scaleY)
        {
            float k = width / height;
            PointF offset = new PointF();
            if (rect.Width / rect.Height > k)
            {
                float contentWidth = k * rect.Height;
                k = rect.Height / height;
                offset.X = (rect.Width - contentWidth) / 2;
            }
            else
            {
                float contentHeight = rect.Width / k;
                k = rect.Width / width;
                offset.Y = (rect.Height - contentHeight) / 2;
            }

            scaleX = k;
            scaleY = k;
            offsX = offset.X + rect.X;
            offsY = offset.Y + rect.Y;
        }

        private void DrawContent(
            GcDXGraphics graphics,
            D2D.Device device,
            ref GcD2DBitmap bitmap,
            int pageIndex,
            int landscapeAngle,
            float printerDpi,
            SizeF paperSize,
            RectangleF printableArea,
            RenderingCache renderingCache,
            FontCache fontCache)
        {
            Page page = _doc.Pages[pageIndex];

            SizeF pageSize = page.GetRenderSize(graphics.Resolution, graphics.Resolution);

            if (PageScaling == PageScaling.FitToPaper)
                printableArea = new RectangleF(0, 0, paperSize.Width, paperSize.Height);

            RectangleF alignRect;
            int rotationAngle;
            if (AutoRotate && ShouldRotate(printableArea.Size, pageSize))
            {
                rotationAngle = landscapeAngle;
                alignRect = new RectangleF(0, 0, printableArea.Height, printableArea.Width);
            }
            else
            {
                rotationAngle = 0;
                alignRect = new RectangleF(0, 0, printableArea.Width, printableArea.Height);
            }
            AlignInRect(alignRect,
                pageSize.Width,
                pageSize.Height,
                out float offsX,
                out float offsY,
                out float scaleX,
                out float scaleY);
            Matrix3x2 m;
            switch (rotationAngle)
            {
                case 0:
                    m = new Matrix3x2(scaleX, 0, 0, scaleY, printableArea.X + offsX, printableArea.Y + offsY);
                    break;
                case 90:
                    m = Matrix3x2.Multiply(Matrix3x2.CreateScale(scaleX, scaleY),
                        Matrix3x2.CreateRotation((float)(90 * Math.PI / 180)));
                    m.M31 = offsY + pageSize.Height * scaleY + printableArea.X;
                    m.M32 = printableArea.Y;
                    break;
                case 270:
                    m = Matrix3x2.Multiply(Matrix3x2.CreateScale(scaleX, scaleY),
                        Matrix3x2.CreateRotation((float)(270 * Math.PI / 180)));
                    m.M31 = printableArea.X;
                    m.M32 = offsX + pageSize.Width * scaleX + printableArea.Y;
                    break;
                default:
                    throw new ArgumentOutOfRangeException("Valid values are 90 & 270.");
            }

            TransparencyFeatures tf = page.GetTransparencyFeatures();
            if (tf == TransparencyFeatures.None)
            {
                // the page's content stream does not use transparency features, render directly on the graphics:
                graphics.Transform = m;
                page.Draw(graphics, new RectangleF(0, 0, pageSize.Width, pageSize.Height), true, true, renderingCache, true);
            }
            else
            {
                // use rendering via GcD2DGraphics:
                if (bitmap == null)
                {
                    bitmap = new GcD2DBitmap(device, graphics.Factory);
                    bitmap.SetFontCache(fontCache);
                }
                SizeF pageSizeScaled = new SizeF(
                    pageSize.Width * scaleX * printerDpi / 96f,
                    pageSize.Height * scaleY * printerDpi / 96f);
                Size bitmapSize = new Size(
                    (int)(pageSizeScaled.Width + 0.5f),
                    (int)(pageSizeScaled.Height + 0.5f));
                if (bitmap.PixelWidth != bitmapSize.Width || bitmap.PixelHeight != bitmapSize.Height)
                    bitmap.CreateImage(bitmapSize.Width, bitmapSize.Height, printerDpi, printerDpi);
                using (GcD2DBitmapGraphics bg = bitmap.CreateGraphics(Color.FromArgb(0)))
                {
                    page.Draw(bg, new RectangleF(0, 0, pageSizeScaled.Width, pageSizeScaled.Height), true, true, renderingCache, true);
                }
                // draw bitmap on graphics with offset:
                graphics.Transform = m;
                graphics.RenderTarget.DrawBitmap(bitmap.Bitmap, new DX.RectF(offsX, offsY, pageSize.Width, pageSize.Height));
            }
        }

        private bool PrintPage(
            string printerName,
            D2D.DeviceContext rt,
            D2D.PrintControl printControl,
            D2D.Device device,
            GcDXGraphics graphics,
            int pageIndex,
            int landscapeAngle,
            float printerDpi,
            float printerPageWidthPx,
            float printerPageHeightPx,
            RectangleF printableArea,
            int pageCount,
            RenderingCache renderingCache,
            FontCache fontCache,
            ref int pageNo,
            ref GcD2DBitmap bitmap)
        {
            D2D.CommandList printCommandList = D2D.CommandList.Create(rt);
            rt.SetTarget(printCommandList);
            rt.BeginDraw();

            DrawContent(
                graphics,
                device,
                ref bitmap,
                pageIndex,
                landscapeAngle,
                printerDpi,
                new SizeF(printerPageWidthPx, printerPageHeightPx),
                printableArea,
                renderingCache,
                fontCache);

            bool res = rt.EndDraw(true);
            if (res)
            {
                printCommandList.Close();
                printControl.AddPage(printCommandList, new DX.Size2F(printerPageWidthPx, printerPageHeightPx));
            }
            printCommandList.Dispose();

            //
            pageNo++;
            if (!OnLongOperation(0.2 + ((double)pageNo / (double)pageCount) * 0.8, true))
                return false;

            if (!res)
                throw new Exception($"Error while printing on printer [{printerName}].");

            return true;
        }

        private unsafe void Print(
            string printerName,
            STG.ComStream jobPrintTicketStream,
            float printerPageWidthPx,
            float printerPageHeightPx,
            RectangleF printableArea,
            int dpi,
            bool autoCollate,
            int copies,
            bool collate,
            OutputRange outputRange)
        {
            DXGI.PrintDocumentPackageTargetFactory documentTargetFactory = null;
            DXGI.PrintDocumentPackageTarget documentTarget = null;
            D2D.Factory1 d2dFactory = null;
            D3D.DeviceContext d3dContext = null;
            D3D.Device d3dDevice = null;
            DXGI.Device1 dxgiDevice = null;
            D2D.Device d2dDevice = null;
            D2D.PrintControl printControl = null;
            WIC.ImagingFactory2 wicFactory = null;
            D2D.DeviceContext rt = null;
            DW.Factory1 dwFactory = null;
            GcD2DBitmap bitmap = null;
            RenderingCache renderingCache = _renderingCache;
            if (renderingCache == null)
                renderingCache = new RenderingCache();
            try
            {
#if DEBUG && false
                jobPrintTicketStream.Seek(0, System.IO.SeekOrigin.Begin);
                var tfile = Path.GetTempFileName();
                using (var fs = new FileStream(tfile, System.IO.FileMode.Create))
                {
                    byte[] data = new byte[1024 * 16];
                    fixed (byte* d = data)
                    {
                        while (true)
                        {
                            var read = jobPrintTicketStream.Read((IntPtr)d, data.Length);
                            if (read <= 0)
                                break;
                            fs.Write(data, 0, data.Length);
                        }
                    }
                }
#endif
                //
                string printJobName = _printJobName;
                if (string.IsNullOrEmpty(printJobName))
                    printJobName = "GcPdf Print Job";

                //
                documentTargetFactory = DXGI.PrintDocumentPackageTargetFactory.Create();
                try
                {
                    documentTarget = documentTargetFactory.CreateDocumentPackageTargetForPrintJob(
                        printerName,
                        printJobName,
                        jobPrintTicketStream);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Cannot create print job for printer [{0}].\r\nException:\r\n{1}", printerName, ex.Message));
                }
                finally
                {
                    documentTargetFactory.Dispose();
                    documentTargetFactory = null;
                }
                if (documentTarget == null)
                    // null is returned only if user cancels printing
                    throw new OperationCanceledException();

                // initialize device resources
                d2dFactory = D2D.Factory1.Create(D2D.FactoryType.SingleThreaded);
                wicFactory = WIC.ImagingFactory2.Create();
                dwFactory = DW.Factory1.Create(DW.FactoryType.Shared);
                D3D.FeatureLevel[] featureLevels = new D3D.FeatureLevel[]
                {
                    D3D.FeatureLevel.Level_11_1,
                    D3D.FeatureLevel.Level_11_0,
                    D3D.FeatureLevel.Level_10_1,
                    D3D.FeatureLevel.Level_10_0
                };
                D3D.FeatureLevel actualLevel;
                d3dContext = null;
                d3dDevice = new D3D.Device(IntPtr.Zero);

                DX.HResult result = DX.HResult.Ok;
                for (int i = 0; i <= 1; i++)
                {
                    // use WARP if hardware is not available
                    D3D.DriverType driverType = i == 0 ? D3D.DriverType.Hardware : D3D.DriverType.Warp;
                    result = D3D.D3D11.CreateDevice(null, driverType, IntPtr.Zero, D3D.DeviceCreationFlags.BgraSupport | D3D.DeviceCreationFlags.SingleThreaded,
                        featureLevels, featureLevels.Length, D3D.D3D11.SdkVersion, d3dDevice, out actualLevel, out d3dContext);
                    if (result.Code != unchecked((int)0x887A0004)) // DXGI_ERROR_UNSUPPORTED
                    {
                        break;
                    }
                }
                result.CheckError();
                //
                dxgiDevice = d3dDevice.QueryInterface<DXGI.Device1>();
                d3dContext.Dispose();
                d3dContext = null;
                //
                d2dDevice = d2dFactory.CreateDevice(dxgiDevice);
                //
                D2D.PrintControlProperties printControlProperties = new D2D.PrintControlProperties
                {
                    FontSubset = D2D.PrintFontSubsetMode.Default,
                    RasterDPI = (float)dpi,
                    ColorSpace = D2D.ColorSpace.SRgb
                };
                //
                if (!OnLongOperation(0.2, true))
                    throw new OperationCanceledException();
                //
                printControl = D2D.PrintControl.Create(d2dDevice, wicFactory, documentTarget.NativePointer, printControlProperties);
                if (printControl == null)
                    throw new OperationCanceledException();
                rt = D2D.DeviceContext.Create(d2dDevice, D2D.DeviceContextOptions.None);
                //
                int totalPageCount = _doc.Pages.Count;
                int pageNo = 0;
                int pageCount = outputRange.GetPageCount(1, totalPageCount) * copies;
                int landscapeAngle = PrinterSettings.LandscapeAngle;
                using (FontCache fontCache = new FontCache(dwFactory))
                using (GlyphPathCache glyphPathCache = new GlyphPathCache())
                using (D2D.SolidColorBrush brush = rt.CreateSolidColorBrush(DX.ColorF.Black, null))
                using (GcDXGraphics graphics = new GcDXGraphics(rt, d2dFactory, null, fontCache, brush, glyphPathCache, false))
                {
                    IEnumerator<int> pages = outputRange.GetEnumerator(1, totalPageCount);
                    if (autoCollate)
                    {
                        if (collate)
                        {
                            for (int i = 0; i < copies; i++)
                            {
                                pages.Reset();
                                while (pages.MoveNext())
                                {
                                    PrintPage(
                                        printerName,
                                        rt,
                                        printControl,
                                        d2dDevice,
                                        graphics,
                                        pages.Current - 1,
                                        landscapeAngle,
                                        dpi,
                                        printerPageWidthPx,
                                        printerPageHeightPx,
                                        printableArea,
                                        pageCount,
                                        renderingCache,
                                        fontCache,
                                        ref pageNo,
                                        ref bitmap);
                                }
                            }
                        }
                        else
                        {
                            while (pages.MoveNext())
                            {
                                for (int i = 0; i < copies; i++)
                                {
                                    PrintPage(
                                        printerName,
                                        rt,
                                        printControl,
                                        d2dDevice,
                                        graphics,
                                        pages.Current - 1,
                                        landscapeAngle,
                                        dpi,
                                        printerPageWidthPx,
                                        printerPageHeightPx,
                                        printableArea,
                                        pageCount,
                                        renderingCache,
                                        fontCache,
                                        ref pageNo,
                                        ref bitmap);
                                }
                            }
                        }
                    }
                    else
                    {
                        while (pages.MoveNext())
                        {
                            PrintPage(
                                printerName,
                                rt,
                                printControl,
                                d2dDevice,
                                graphics,
                                pages.Current - 1,
                                landscapeAngle,
                                dpi,
                                printerPageWidthPx,
                                printerPageHeightPx,
                                printableArea,
                                pageCount,
                                renderingCache,
                                fontCache,
                                ref pageNo,
                                ref bitmap);
                        }
                    }
                }
                rt.Dispose();
                rt = null;
                printControl.Close();
            }
            finally
            {
                if (bitmap != null)
                    bitmap.Dispose();
                if (rt != null)
                    rt.Dispose();
                if (printControl != null)
                    printControl.Dispose();
                if (d2dDevice != null)
                    d2dDevice.Dispose();
                if (dxgiDevice != null)
                    dxgiDevice.Dispose();
                if (d3dContext != null)
                    d3dContext.Dispose();
                if (d3dDevice != null)
                    d3dDevice.Dispose();
                if (dwFactory != null)
                    dwFactory.Dispose();
                if (wicFactory != null)
                    wicFactory.Dispose();
                if (d2dFactory != null)
                    d2dFactory.Dispose();
                if (documentTarget != null)
                    documentTarget.Dispose();
                if (documentTargetFactory != null)
                    documentTargetFactory.Dispose();
                if (_renderingCache == null)
                    renderingCache.Dispose();
            }
        }
        #endregion

        #region Events
        public event Func<double, bool, bool>? LongOperation;
        #endregion

        #region pinvoke
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class DEVMODE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
            public string? dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public short dmOrientation;
            public short dmPaperSize;
            public short dmPaperLength;
            public short dmPaperWidth;
            public short dmScale;
            public short dmCopies;
            public short dmDefaultSource;
            public short dmPrintQuality;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 0x20)]
            public string? dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmICCManufacturer;
            public int dmICCModel;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }

        /// <summary>
        /// Fields of DEVMODE structure.
        /// </summary>
        private class DM
        {
            public const int
                DM_ORIENTATION = 0x00001,
                DM_PAPERSIZE = 2,
                DM_PAPERLENGTH = 4,
                DM_PAPERWIDTH = 8,
                DM_COPIES = 0x100,
                DM_DEFAULTSOURCE = 0x200,
                DM_PRINTQUALITY = 0x400,
                DM_COLOR = 0x800,
                DM_YRESOLUTION = 0x2000,
                DM_COLLATE = 0x00008000,
                DM_BITSPERPEL = 0x40000,
                DM_PELSWIDTH = 0x80000,
                DM_PELSHEIGHT = 0x100000,
                DM_DISPLAYFREQUENCY = 0x400000;
        }

        [DllImport("kernel32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern IntPtr GlobalLock(IntPtr handle);

        [DllImport("kernel32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern bool GlobalUnlock(IntPtr handle);

        [DllImport("kernel32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern IntPtr GlobalFree(IntPtr hMem);

        [DllImport("ole32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern int CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, ref IntPtr istream);

        [DllImport("prntvpt.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern int PTConvertDevModeToPrintTicket(IntPtr hProvider, int cbDevmode, IntPtr devMode, int scope, IntPtr printTicket);

        [DllImport("prntvpt.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int PTOpenProvider(string pszPrinterName, int dwVersion, ref IntPtr phProvider);

        [DllImport("prntvpt.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern int PTCloseProvider(IntPtr hProvider);
        #endregion
    }

}