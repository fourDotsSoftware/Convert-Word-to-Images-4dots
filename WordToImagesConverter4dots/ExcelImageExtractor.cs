using System;
using System.Collections.Generic;

using System.Text;

using System.Drawing;
using System.Reflection;
using System.Threading;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;
using System.Windows.Media.Imaging;

namespace WordToImagesConverter4dots
{
    public class ExcelImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();

        public List<FromToWordImage> ExtractedFromToWordImages = new List<FromToWordImage>();

        public string err = "";

        public bool ExtractImages(string filepath)
        {
            err = "";

            Image image = null;
            object WordAppSelection = null;
            object HeaderRangeShape = null;
            int iHeaderRangeShapesCount = -1;
            object HeaderRangeShapesCount = null;
            object HeaderRangeShapes = null;
            object HeaderRange = null;
            object Header = null;
            object oDocuments = null;
            object doc = null;
            object Sections = null;

            object pImage = null;
            object pImageImage = null;

            try
            {
                OfficeHelper.CreateExcelApplication();

                oDocuments = OfficeHelper.ExcelApp.GetType().InvokeMember("Workbooks", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.ExcelApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                //System.Threading.Thread.Sleep(100);

                /*
                OfficeHelper.ExcelApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.ExcelApp, null);
                */
                System.Threading.Thread.Sleep(200);

                Sections = doc.GetType().InvokeMember("Sheets", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                object SectionsCount = Sections.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Sections, null);
                int iSectionsCount = (int)SectionsCount;

                for (int m1 = 1; m1 <= iSectionsCount; m1++)
                {
                    object oSlide = doc.GetType().InvokeMember("Sheets", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, new object[] { m1 });

                    object oShapes = oSlide.GetType().InvokeMember("Pictures", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oSlide, null);

                    object oShapesCount = oShapes.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oShapes, null);

                    int iShapesCount = (int)oShapesCount;

                    for (int m2 = 1; m2 <= iShapesCount; m2++)
                    {
                        object oShape = oSlide.GetType().InvokeMember("Pictures", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oSlide, new object[] { m2 });

                        string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory);

                        //object[] oParam = new object[] { imgfp, 2, 0, 0, 4 };

                        object[] oParam = new object[] { 1, 2 };

                        oShape.GetType().InvokeMember("CopyPicture", BindingFlags.InvokeMethod, null, oShape, oParam);

                        //shape.Export("C:\\myPathtoFile", Microsoft.Office.Interop.PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0, Microsoft.Office.Interop.PowerPoint.PpExportMode.ppScaleXY)

                        FromToWordImage wim = new FromToWordImage();                        
                        wim.WordFilepath = filepath;
                        
                        Thread thread = new Thread(new ParameterizedThreadStart(SaveInlineShape));
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start(wim);
                        thread.Join();

                        oShape = null;

                    }

                    oSlide = null;
                    oShapes = null;
                    oShapesCount = null;
                }

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Replace Image for Document") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }

        Bitmap GetBitmap(BitmapSource source)
        {
            Bitmap bmp = new Bitmap(
              source.PixelWidth,
              source.PixelHeight,
              PixelFormat.Format32bppPArgb);
            BitmapData data = bmp.LockBits(
              new Rectangle(System.Drawing.Point.Empty, bmp.Size),
              ImageLockMode.WriteOnly,
              PixelFormat.Format32bppPArgb);
            source.CopyPixels(
              Int32Rect.Empty,
              data.Scan0,
              data.Height * data.Stride,
              data.Stride);
            bmp.UnlockBits(data);
            return bmp;
        }

        protected void SaveInlineShape(object owim)
        {
            try
            {
                if (System.Windows.Clipboard.GetDataObject() != null)
                {
                    System.Windows.IDataObject data = System.Windows.Clipboard.GetDataObject();
                    if (data.GetDataPresent(System.Windows.DataFormats.Bitmap))
                    {
                        System.Windows.Interop.InteropBitmap image = (System.Windows.Interop.InteropBitmap)data.GetData(System.Windows.DataFormats.Bitmap, true);

                        Bitmap bmp = GetBitmap(image);

                        //string imgfp = System.IO.Path.Combine(Module.CurrentImagesDirectory, Guid.NewGuid().ToString() + ".bmp");

                        FromToWordImage wim = owim as FromToWordImage;

                        string imgfp = frmOptions.GetSaveFilepath(wim.WordFilepath, Module.CurrentImagesDirectory);

                        ExtractedFilepaths.Add(imgfp);

                        //bmp.Save(imgfp);

                        frmOptions.SaveImage(imgfp, bmp);

                        wim.ImageFilepath = imgfp;

                        ExtractedFromToWordImages.Add(wim);

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imgfp;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
