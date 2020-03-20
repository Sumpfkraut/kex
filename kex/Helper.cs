using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;
using System.Text;

namespace kex
{
    public class Helper
    {
        public string RemoveSpecialCharacters(string str)
        {
            str = str.Replace(" ", "-");
            str = str.Replace("_", "-");
            str = str.Replace("+", "-");
            str = str.Replace("&", "-");
            str = str.Replace("Ö", "OE");
            str = str.Replace("Ä", "AE");
            str = str.Replace("Ü", "UE");
            str = str.Replace("É", "E");
            str = str.Replace("À", "A");
            str = str.Replace("È", "E");
            str = str.Replace("°", "GRD");

            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_' || c == '-')
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        public void Zoom(Editor ed, Extents3d ext)
        {
            if (ed == null)
                throw new ArgumentNullException("ed");
            using (ViewTableRecord view = ed.GetCurrentView())
            {
                Matrix3d worldToEye = Matrix3d.WorldToPlane(view.ViewDirection) *
                    Matrix3d.Displacement(Point3d.Origin - view.Target) *
                    Matrix3d.Rotation(view.ViewTwist, view.ViewDirection, view.Target);
                ext.TransformBy(worldToEye);
                view.Width = ext.MaxPoint.X - ext.MinPoint.X;
                view.Height = ext.MaxPoint.Y - ext.MinPoint.Y;
                view.CenterPoint = new Point2d(
                    (ext.MaxPoint.X + ext.MinPoint.X) / 2.0,
                    (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0);
                ed.SetCurrentView(view);
            }
        }

        public void ZoomExtents(Document doc)
        {
            Database db = doc.Database;
            db.UpdateExt(false);
            Extents3d ext = (short)Application.GetSystemVariable("cvport") == 1 ?
                new Extents3d(db.Pextmin, db.Pextmax) :
                new Extents3d(db.Extmin, db.Extmax);
            Zoom(doc.Editor, ext);
        }

        public bool OpenDirectory(string directory)
        {
            try { System.Diagnostics.Process.Start(directory); }
            catch { return false; }
            return true;
        }

        public bool CheckDirectory(string directory, bool createIfNotExist, bool deleteExistingFiles, string deleteFilesExtension = "")
        {
            if (!Directory.Exists(directory))
            {
                if (createIfNotExist)
                {
                    try { Directory.CreateDirectory(directory); }
                    catch { return false; }
                }
                else
                {
                    return false;
                }
            }

            if (deleteExistingFiles)
            {
                DirectoryInfo dir = new DirectoryInfo(directory);
                foreach (FileInfo file in dir.GetFiles())
                {
                    if (deleteFilesExtension != "")
                    {
                        if (file.Extension.ToUpper() == deleteFilesExtension)
                        {
                            try { file.Delete(); }
                            catch { return false; }
                        }
                    }
                    else
                    {
                        try { file.Delete(); }
                        catch { return false; }
                    }
                }
            }
            return true;
        }
    }
}
