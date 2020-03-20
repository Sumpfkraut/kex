using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Windows;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace kex
{
    public class Kex : IExtensionApplication
    {
        private static Autodesk.AutoCAD.Windows.PaletteSet paletteSetKex = null;
        KexControl _kexControl = new KexControl();

        public void Initialize()
        {
            acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Cscript kex.dll is loaded.." + Environment.NewLine);
        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }

        [CommandMethod("KEX")]
        public void StlEx()
        {
            if (paletteSetKex == null)
            {
                try
                {
                    paletteSetKex = new PaletteSet("Kex - STL export", new Guid("8ba5e63b-34ad-4dd2-a9aa-85ad686589ef"));
                    paletteSetKex.Add("Kex - STL export", _kexControl);
                    //paletteSetKex.KeepFocus = true;
                    paletteSetKex.Style = PaletteSetStyles.NameEditable | PaletteSetStyles.ShowCloseButton;
                    paletteSetKex.StateChanged += new PaletteSetStateEventHandler(PaletteSetKex2_StateChanged);
                    paletteSetKex.Location = new System.Drawing.Point(200, 300);
                    paletteSetKex.MinimumSize = new System.Drawing.Size(502, 902);
                    paletteSetKex.Size = new System.Drawing.Size(502, 902);
                    paletteSetKex.Visible = true;
                    paletteSetKex.DockEnabled = DockSides.None;
                    paletteSetKex.Dock = DockSides.None;
                    _kexControl.LoadSettings();
                }
                catch
                {
                    MessageBox.Show("Failed to create the palette.");
                }
            }
            else
            {
                try
                {
                    paletteSetKex.Visible = true;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Original error: " + ex.Message);
                }
            }
        }

        public void PaletteSetKex2_StateChanged(object sender, PaletteSetStateEventArgs e)
        {
            //Editor ed = acApp.DocumentManager.MdiActiveDocument.Editor;
            //ed.WriteMessage("\nPalette StateChanged ! New State is : " + e.NewState.ToString()); //Hide or Show
            if (paletteSetKex != null && e.NewState.ToString() == "Show")
            {
                _kexControl.LoadSettings();
            }
        }

        [CommandMethod("RL")]
        public void RandomLayer()
        {
            DocumentCollection documentManager = acApp.DocumentManager;
            Document document = documentManager.MdiActiveDocument;
            Editor editor = document.Editor;
            Database database = document.Database;
            int counter = 1;
            bool displayErrorMessage = true;

            using (Transaction tr = database.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(database.LayerTableId, OpenMode.ForRead);
                PromptSelectionResult psr = editor.GetSelection();
                if(psr.Status == PromptStatus.OK)
                {
                    SelectionSet sset = psr.Value;
                    foreach(SelectedObject so in sset)
                    {
                        if(so != null)
                        {
                            try
                            {
                                Entity ent = (Entity)tr.GetObject(so.ObjectId, OpenMode.ForWrite);
                                if (ent != null)
                                {
                                    string myNewLayer = "000_" + counter + "_" + ent.Layer.ToString();
                                    for (int i = 0; i < 10000; i++)
                                    {
                                        if (lt.Has(myNewLayer))
                                        {
                                            counter++;
                                            myNewLayer = "000_" + counter + "_" + ent.Layer.ToString();
                                        }
                                        else
                                        {
                                            lt.UpgradeOpen();
                                            LayerTableRecord ltr = new LayerTableRecord();
                                            ltr.Name = myNewLayer;
                                            lt.Add(ltr);
                                            tr.AddNewlyCreatedDBObject(ltr, true);
                                            break;
                                        }
                                    }
                                    try
                                    {
                                        ent.Layer = myNewLayer;
                                        counter++;
                                    }
                                    catch
                                    {
                                        if (displayErrorMessage)
                                        {
                                            displayErrorMessage = false;
                                            MessageBox.Show("Can't apply the new layer.");
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                if (displayErrorMessage)
                                {
                                    displayErrorMessage = false;
                                    MessageBox.Show("One or more objects are on locked or freezed layers.");
                                }
                            }
                        }
                    }
                }
                tr.Commit();
            }
        }
    }
}