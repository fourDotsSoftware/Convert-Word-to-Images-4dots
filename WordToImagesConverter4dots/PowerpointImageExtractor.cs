using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Reflection;

namespace WordToImagesConverter4dots
{
    public class PowerpointImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();        

        public string err = "";

        public bool ExtractImages(string filepath,string slideranges)
        {
            err = "";
            
            object oDocuments = null;
            object doc = null;
            object Sections = null;            

            try
            {
                OfficeHelper.CreatePowerPointApplication();

                oDocuments = OfficeHelper.PPApp.GetType().InvokeMember("Presentations", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.PPApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                /*
                System.Threading.Thread.Sleep(100);

                OfficeHelper.PPApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.PPApp, null);
                */

                System.Threading.Thread.Sleep(200);

                Sections = doc.GetType().InvokeMember("Slides", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                object SectionsCount = Sections.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Sections, null);
                int iSectionsCount = (int)SectionsCount;

                StringRange stringrange = new StringRange(slideranges);

                for (int m1 = 1; m1 <= iSectionsCount; m1++)
                {
                    if (stringrange.IsInRange(m1))
                    {
                        object oSlide = doc.GetType().InvokeMember("Slides", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, new object[] { m1 });

                        string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory,m1);

                        object[] oParam = new object[] { imgfp, frmOptions.GetImageExtension().ToUpper() };

                        oSlide.GetType().InvokeMember("Export", BindingFlags.InvokeMethod, null, oSlide, oParam);

                        oSlide = null;

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imgfp;
                        }
                    }
                }

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Powerpoint to Images") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }                
    }
}
