using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using acApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections;
using System.Xml;
using System.Threading.Tasks;

namespace kex
{
    public partial class KexControl : UserControl
    {
        private readonly Helper _helper = new Helper();
        private SortedDictionary<string, string> _customPropertiesDict = new SortedDictionary<string, string>();
        private string _layerSeperator = "_";
        private int _progressbarSteps = 3;
        private double _progress = 0;
        private bool _running = false;

        public KexControl()
        {
            InitializeComponent();
            LoadProperties();
        }

        public void LoadProperties()
        {
            //Options Tab
            textBox3.Text = Properties.Settings.Default.KEX_STL_PATH;
            textBox5.Text = Environment.UserName;
            textBox4.Text = Properties.Settings.Default.KEX_QUALITY.ToString();
            checkBox5.Checked = Properties.Settings.Default.KEX_RESET_UCS;
            checkBox6.Checked = Properties.Settings.Default.KEX_MOVE_XYZ;
            checkBox8.Checked = Properties.Settings.Default.KEX_USE_NEW_DRAWING;
            checkBox9.Checked = Properties.Settings.Default.KEX_CLOSE_NEW_DRAWING;
            textBox6.Text = Properties.Settings.Default.KEX_GROUP_PREFIX;
            textBox7.Text = Properties.Settings.Default.KEX_C4D_SUB_FOLDER;
            textBox8.Text = Properties.Settings.Default.KEX_RENDER_SUB_FOLDER;

            //Advanced
            checkBox10.Checked = Properties.Settings.Default.KEX_USE_TIMEOUT;

            //Export Tab
            checkBox1.Checked = Properties.Settings.Default.KEX_EXPLODE;
            checkBox2.Checked = Properties.Settings.Default.KEX_EXPLODE_NESTED;
            checkBox7.Checked = Properties.Settings.Default.KEX_SOLID_OUTSIDE_BLOCK;
            checkBox4.Checked = Properties.Settings.Default.KEX_EXPORT_HIDDEN;
            checkBox3.Checked = Properties.Settings.Default.KEX_USE_EXCEPTION_LAYERS;
            textBox2.Enabled = Properties.Settings.Default.KEX_USE_EXCEPTION_LAYERS;
            textBox2.Text = Properties.Settings.Default.KEX_EXCEPTION_LAYERS;
            button1.Enabled = true;
            button5.Enabled = false;
        }

        public void SaveProperties()
        {
            //Options Tab
            Properties.Settings.Default.KEX_STL_PATH = textBox3.Text;
            bool isDouble = double.TryParse(textBox4.Text, out double quali);
            if (isDouble && quali > 0 && quali <= 10)
                Properties.Settings.Default.KEX_QUALITY = quali;

            Properties.Settings.Default.KEX_RESET_UCS = checkBox5.Checked;
            Properties.Settings.Default.KEX_MOVE_XYZ = checkBox6.Checked;
            Properties.Settings.Default.KEX_USE_NEW_DRAWING = checkBox8.Checked;
            Properties.Settings.Default.KEX_CLOSE_NEW_DRAWING = checkBox9.Checked;
            Properties.Settings.Default.KEX_GROUP_PREFIX = textBox6.Text;
            Properties.Settings.Default.KEX_C4D_SUB_FOLDER = textBox7.Text;
            Properties.Settings.Default.KEX_RENDER_SUB_FOLDER = textBox8.Text;

            //Export Tab
            Properties.Settings.Default.KEX_EXPLODE = checkBox1.Checked;
            Properties.Settings.Default.KEX_EXPLODE_NESTED = checkBox2.Checked;
            Properties.Settings.Default.KEX_SOLID_OUTSIDE_BLOCK = checkBox7.Checked;
            Properties.Settings.Default.KEX_EXPORT_HIDDEN = checkBox4.Checked;
            Properties.Settings.Default.KEX_USE_EXCEPTION_LAYERS = checkBox3.Checked;
            Properties.Settings.Default.KEX_EXCEPTION_LAYERS = textBox2.Text;

            Properties.Settings.Default.Save();
        }


        public void LoadSettings()
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            GetCustomProperties();
            if (_customPropertiesDict.Count > 0)
            {
                foreach (var item in _customPropertiesDict)
                {
                    comboBox1.Items.Add(item.Key);
                    comboBox2.Items.Add(item.Key);
                }

                //C4D root object name
                if(_customPropertiesDict.Keys.Contains(Properties.Settings.Default.KEX_ROOT_PROPERTY_NAME))
                {
                    comboBox1.SelectedIndex = comboBox1.FindStringExact(Properties.Settings.Default.KEX_ROOT_PROPERTY_NAME);
                }

                //Render output path
                if (_customPropertiesDict.Keys.Contains(Properties.Settings.Default.KEX_RENDERPATH_PROPERTY_NAME))
                {
                    comboBox2.SelectedIndex = comboBox2.FindStringExact(Properties.Settings.Default.KEX_RENDERPATH_PROPERTY_NAME);
                }
            }
        }

        public void GetCustomProperties()
        {
            IDictionaryEnumerator cp = acApp.DocumentManager.MdiActiveDocument.Database.SummaryInfo.CustomProperties;
            cp.Reset();
            _customPropertiesDict.Clear();
            while (cp.MoveNext())
            {
                _customPropertiesDict.Add(cp.Key.ToString(), cp.Value.ToString());
            }
        }

        //EXPORT BUTTON
        private void Button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button5.Enabled = true;
            textBox1.Clear();
            SaveProperties();
            _progress = 0;
            progressBar1.Value = progressBar1.Minimum;
            groupBox4.BackColor = System.Drawing.Color.OrangeRed;
            groupBox2.Text = "Results";
            label10.Text = "Blocks...";
            label11.Text = "Solids...";

            label9.Text = "Create a new drawing and copy blocks/solids...";
            label9.Update();
            label9.Refresh();
            this.Refresh();

            DocumentCollection documentManager = acApp.DocumentManager;
            Document doc = documentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Document originalDocument = documentManager.MdiActiveDocument;

            List<Tuple<string, string, string>> blocknameForXml = new List<Tuple<string, string, string>>();

            //Check the target directory for STL files
            string targetDirectory = Path.Combine(Properties.Settings.Default.KEX_STL_PATH, Environment.UserName);
            bool targetDirectoryOk = _helper.CheckDirectory(targetDirectory, true, true, ".STL");
            if (!targetDirectoryOk)
            {
                ed.WriteMessage("Check target directory failed: " + targetDirectory + " -> Can't create it, or can't delete existing files!\n");
                return;
            }
            else
            {
                ed.WriteMessage("Check target directory done: " + targetDirectory + " -> Ready to write files here...\n");
            }

            //Save quality settings
            bool isdouble = double.TryParse(textBox4.Text, out double quali);
            if(isdouble && quali > 0 && quali <= 10 && quali != Properties.Settings.Default.KEX_QUALITY)
            {
                SaveFacetresValue(quali);
            }

            //Get exception layers
            List<string> exceptionLayerNames = new List<string>();
            if (checkBox3.Checked && textBox2.Text.Length > 0)
            {
                try
                {
                    exceptionLayerNames = textBox2.Text.Split(',').ToList();
                }
                catch (System.Exception acadEx)
                {
                    ed.WriteMessage("Split exception layers failed: " + textBox2.Text + "\n -> Original error: " + acadEx.Message + "\n");
                }
            }

            //Get and set system variables
            double originalfacetres = System.Convert.ToDouble(acApp.GetSystemVariable("FACETRES"));
            ed.WriteMessage("FACETRES current value: " + originalfacetres.ToString() + "\n");
            acApp.SetSystemVariable("FILEDIA", 0);
            acApp.SetSystemVariable("FACETRES", Properties.Settings.Default.KEX_QUALITY);

            //Reset UCS if checkbox is checked
            if (checkBox5.Checked)
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
            }

            //Get drawing infos
            Dictionary<string, string> drawingInfos = GetDrawingInfos(ed);
            ed.WriteMessage("C4D path: " + drawingInfos["cadPath"] +
                "\nCAD Name: " + drawingInfos["cadName"] +
                "\nRender path: " + drawingInfos["pdfPath"] + 
                "\nC4D filename: " + drawingInfos["c4dFileName"] + 
                "\nCustom Property: " + Properties.Settings.Default.KEX_ROOT_PROPERTY_NAME + " -> " +  drawingInfos["customProperty"] + 
                "\n");

            //If use a new drawing is checked
            if (Properties.Settings.Default.KEX_USE_NEW_DRAWING)
            {
                //overwrite the document, database and editor variable with the new drawing
                doc = CreateNewDrawing(documentManager);
                documentManager.MdiActiveDocument = doc;
                db = doc.Database;
                ed = doc.Editor;
            }

            //Get all BlockReference and Solid3d
            List<BlockReference> brefList = new List<BlockReference>();
            List<Solid3d> solidList = new List<Solid3d>();

            //Lock the document
            using (DocumentLock acDocLock = doc.LockDocument())
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    label9.Text = "Get objects and move them...";
                    label9.Update();
                    label9.Refresh();
                    this.Refresh();
                    double progressStep0 = ((double)progressBar1.Maximum / _progressbarSteps) / 3;
                    progressBar1.Value = UpdateProgressBar(progressStep0);
                    progressBar1.Refresh();

                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    BlockTable blckTbl = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord blckTblRcrd = (BlockTableRecord)tr.GetObject(blckTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                    Matrix3d movementMatrix = Matrix3d.Displacement(new Vector3d(10000, 10000, 10000));

                    foreach (ObjectId objectid in blckTblRcrd)
                    {
                        Entity ent = (Entity)tr.GetObject(objectid, OpenMode.ForWrite);

                        foreach (ObjectId layerId in lt)
                        {
                            LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                            if (layer.Name == ent.Layer)
                            {
                                if (!layer.IsLocked && !layer.IsFrozen && !layer.IsOff)
                                {
                                    if (Properties.Settings.Default.KEX_USE_EXCEPTION_LAYERS && exceptionLayerNames.Contains(layer.Name)) break;

                                    if (ent is BlockReference) brefList.Add((BlockReference)ent);
                                    else if (ent is Solid3d && Properties.Settings.Default.KEX_SOLID_OUTSIDE_BLOCK) solidList.Add((Solid3d)ent);

                                    if (checkBox6.Checked)
                                    {
                                        //Move objects if checkbox is checked
                                        try { ent.TransformBy(movementMatrix); }
                                        catch { break; }
                                    }
                                }
                                else
                                {
                                    string entTypeName = ent.GetType().Name;
                                    try
                                    {
                                        ent.Erase();
                                        ed.WriteMessage("Entity (locked, frozen or off) on layer " + layer.Name + " (" + entTypeName + ") erased.\n");
                                    }
                                    catch (System.Exception acadEx)
                                    {
                                        ed.WriteMessage("Erase entity on layer " + layer.Name + " (" + entTypeName + ") failed. -> Original error: " + acadEx.Message + "\n");
                                    }
                                    break;
                                }
                            }
                        }
                    }

                    //Rename Layers
                    label9.Text = "Rename layers...";
                    label9.Update();
                    label9.Refresh();
                    this.Refresh();
                    progressBar1.Value = UpdateProgressBar(progressStep0);
                    progressBar1.Refresh();

                    int layerCounter = 0;
                    foreach (ObjectId layerId in lt)
                    {
                        LayerTableRecord layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForWrite);
                        if (layer.Name != "0") //It's not possible to rename layer 0
                        {
                            string newLayerName = _helper.RemoveSpecialCharacters(layer.Name.ToUpper());
                            try
                            {
                                if (!lt.Has(newLayerName)) layer.Name = newLayerName;
                                else layer.Name = newLayerName + _layerSeperator + layerCounter.ToString();
                            }
                            catch (System.Exception acadEx)
                            {
                                ed.WriteMessage("Rename layer error: " + newLayerName + " -> Original error: " + acadEx.Message + "\n");
                            }
                        }
                        layerCounter++;
                    }
                    progressBar1.Value = UpdateProgressBar(progressStep0);
                    progressBar1.Refresh();

                    //Zoom all
                    _helper.ZoomExtents(doc);
                    ed.WriteMessage("Zoom out...  Total blocks found: " + brefList.Count.ToString() + "... Total solids found: " + solidList.Count.ToString() + "\n");
                    label10.Text = "Blocks: " + brefList.Count.ToString();
                    label11.Text = "Solids: " + solidList.Count.ToString();
                    label10.Refresh();
                    label11.Refresh();
                    this.Refresh();

                    tr.Commit();
                }
            }

            //Refresh viewport
            ed.Regen();
            ed.UpdateScreen();

            using (DocumentLock acDocLock = doc.LockDocument())
            {
                //Total items for the progress bar
                int totalItems = brefList.Count + solidList.Count;

                //Progress bar
                double progressStep1 = ((double)progressBar1.Maximum / _progressbarSteps);
                if (totalItems > 0) progressStep1 /= totalItems;

                //Explode all blocks
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    if (Properties.Settings.Default.KEX_EXPLODE)
                    {
                        groupBox4.BackColor = System.Drawing.Color.DarkOrange;
                        label9.Text = "Explode blocks...";
                        label9.Update();
                        label9.Refresh();
                        this.Refresh();

                        BlockTableRecord modelspace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                        int blockCounter = 0;
                        foreach (BlockReference bref in brefList)
                        {
                            try
                            {
                                //workaround to get the real block (from unnamed blockreferences) to get the real block name -> use IsDynamicBlock direct to the bref can return a wrong value!
                                BlockReference blockRef = tr.GetObject(bref.Id, OpenMode.ForRead) as BlockReference;
                                BlockTableRecord block = null;
                                if (blockRef.IsDynamicBlock) block = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                                else block = tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;

                                string bname = block != null ? block.Name : bref.Name;
                                string newName = _helper.RemoveSpecialCharacters(bname.ToUpper());

                                ExplodeBlocks(tr, db, modelspace, bref.Id, blockCounter, 0, newName, Properties.Settings.Default.KEX_EXPLODE_NESTED, true);
                                blockCounter++;
                            }
                            catch (System.Exception acadEx)
                            {
                                ed.WriteMessage("ExplodeBlocks() error: " + bref.Name + " -> Original error: " + acadEx.Message + "\n");
                            }
                            progressBar1.Value = UpdateProgressBar(progressStep1);
                            progressBar1.Refresh();
                        }
                    }
                    else
                    {
                        progressBar1.Value = UpdateProgressBar(progressStep1 * brefList.Count);
                        progressBar1.Refresh();
                    }

                    //Go for all solid3d
                    if (Properties.Settings.Default.KEX_SOLID_OUTSIDE_BLOCK && solidList.Count > 0)
                    {
                        label9.Text = "Collect solids...";
                        label9.Update();
                        label9.Refresh();
                        this.Refresh();

                        int c = 0;
                        foreach (Solid3d sol in solidList)
                        {
                            Solid3d solid = (Solid3d)tr.GetObject(sol.Id, OpenMode.ForWrite);
                            string newLayerName = sol.Layer + _layerSeperator + c.ToString();
                            try
                            {
                                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                                if (!lt.Has(newLayerName))
                                {
                                    LayerTableRecord newLayer = new LayerTableRecord();
                                    newLayer.Name = newLayerName;
                                    lt.UpgradeOpen();
                                    lt.Add(newLayer);
                                    tr.AddNewlyCreatedDBObject(newLayer, true);
                                }

                                solid.Layer = newLayerName;
                            }
                            catch (System.Exception acadEx)
                            {
                                acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Rename (Solid3d list outside of blocks) Solid3d-layer failed: " + newLayerName + " -> Original error: " + acadEx.Message + "\n");
                            }
                            c++;
                            progressBar1.Value = UpdateProgressBar(progressStep1);
                            progressBar1.Refresh();
                        }
                    }
                    tr.Commit();
                }
            }

            //Refresh viewport
            ed.Regen();
            ed.UpdateScreen();

            using (DocumentLock acDocLock = doc.LockDocument())
            {
                double progressStep2 = ((double)progressBar1.Maximum / _progressbarSteps);
                int layerCount = 0;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (ObjectId obj in lt)
                    {
                        if(!obj.IsErased) layerCount++;
                    }
                    if (layerCount > 0)
                    {
                        progressStep2 /= layerCount;
                        groupBox4.BackColor = System.Drawing.Color.Orange;
                        label9.Text = "Export files...";
                        label9.Update();
                        label9.Refresh();
                        this.Refresh();
                        _running = true;
                        int filecounter = 0;

                        //Get all layers
                        foreach (ObjectId obj in lt)
                        {
                            //Stop the export if the abort button is pressed
                            if (!_running) break;

                            //Get the current layer
                            LayerTableRecord layer = null;
                            try
                            {
                                layer = (LayerTableRecord)tr.GetObject(obj, OpenMode.ForRead);
                            }
                            catch { continue; }

                            try
                            {
                                //Select all items on this layer
                                ObjectIdCollection ents = GetEntitiesOnLayer(ed, layer.Name);
                                if (ents == null || ents.Count <= 0) continue;

                                int c = 0;
                                foreach (ObjectId myId in ents)
                                {
                                    //Prefix for block layers
                                    string prefix = "no-pre-fix";
                                    if(layer.Name.Length > Properties.Settings.Default.KEX_GROUP_PREFIX.Length)
                                        prefix = layer.Name.Substring(0, Properties.Settings.Default.KEX_GROUP_PREFIX.Length);

                                    //Block identification for groups in c4d
                                    string blockIdent = "__SOLIDS";

                                    //Filename for the STL file
                                    string newFileName = layer.Name + _layerSeperator + c.ToString();

                                    if (prefix == Properties.Settings.Default.KEX_GROUP_PREFIX)
                                    {
                                        blockIdent = layer.Name.Remove(0, Properties.Settings.Default.KEX_GROUP_PREFIX.Length).Split('_')[1];
                                        newFileName = newFileName.Remove(0, Properties.Settings.Default.KEX_GROUP_PREFIX.Length);
                                    }

                                    try
                                    {
                                        Solid3d sol = (Solid3d)tr.GetObject(myId, OpenMode.ForRead);
                                        sol.StlOut(Path.Combine(targetDirectory, newFileName + ".stl"), true);
                                        filecounter++;

                                        //await Task.Run(() =>
                                        //{
                                        //    sol.StlOut(Path.Combine(targetDirectory, newFileName + ".stl"), true);
                                        //});

                                    }
                                    catch (System.Exception acadEx)
                                    {
                                        acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("STLOUT failed on layer: " + layer.Name + " -> Original error: " + acadEx.Message + "\n");
                                        continue;
                                    }

                                    //Save block infos for the XML file
                                    blocknameForXml.Add(new Tuple<string, string, string>(blockIdent, newFileName, layer.Name));
                                    c++;

                                    if (Properties.Settings.Default.KEX_USE_TIMEOUT)
                                    {
                                        System.Threading.Thread.Sleep(100);
                                    }
                                }

                            }
                            catch (System.Exception acadEx)
                            {
                                acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Get entities on layer failed: " + layer.Name + " -> Original error: " + acadEx.Message + "\n");
                            }

                            //Update result textbox
                            string newText = layer.Name;
                            if (layer.Name.Length > 48)
                                newText = layer.Name.Substring(0, 48);

                            try { textBox1.Text = newText + Environment.NewLine + textBox1.Text; }
                            catch { textBox1.Text = newText; }
                            textBox1.Update();

                            progressBar1.Value = UpdateProgressBar(progressStep2);
                            progressBar1.Refresh();
                        }
                        _running = false;
                        groupBox2.Text = "STL-files: " + filecounter.ToString();
                    }
                    tr.Commit();
                }

                //Restore system variables
                acApp.SetSystemVariable("FACETRES", originalfacetres);
                acApp.SetSystemVariable("FILEDIA", 1);

                label9.Text = "Write XML file...";
                label9.Update();
                label9.Refresh();
                this.Refresh();

                //Write XML file
                SaveXML(blocknameForXml, drawingInfos);

                //Set progress bar to 100%
                progressBar1.Value = progressBar1.Maximum;
                progressBar1.Refresh();
            }

            //Close the new temporary drawing
            if (Properties.Settings.Default.KEX_USE_NEW_DRAWING && Properties.Settings.Default.KEX_CLOSE_NEW_DRAWING)
            {
                label9.Text = "Close the temporary drawing...";
                label9.Update();
                label9.Refresh();
                this.Refresh();

                doc.CloseAndDiscard();
                documentManager.MdiActiveDocument = originalDocument;
            }

            //Restore buttons and colors
            button1.Enabled = true;
            button5.Enabled = false;
            groupBox4.BackColor = System.Drawing.Color.Transparent;
            label9.Text = "Done!";
            label9.Update();
            label9.Refresh();
            this.Refresh();
        }

        private void SaveXML(List<Tuple<string, string, string>> blocknameForXml, Dictionary<string, string> drawingInfos)
        {
            int xmlElementCounter = 0;
            XmlDocument xmlDocument = new XmlDocument();
            XmlDeclaration xmlDeclaration = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", "no");

            try
            {
                XmlElement rootNode = xmlDocument.CreateElement("RootNode");
                xmlDocument.InsertBefore(xmlDeclaration, xmlDocument.DocumentElement);
                xmlDocument.AppendChild(rootNode);

                foreach (var item in blocknameForXml)
                {
                    try
                    {
                        XmlElement parentNode = xmlDocument.CreateElement("Parent" + xmlElementCounter.ToString());
                        xmlDocument.DocumentElement.PrependChild(parentNode);

                        XmlElement a = xmlDocument.CreateElement("CUSTOM_PROPERTY_VALUE");
                        XmlElement b = xmlDocument.CreateElement("DRAWING_NAME");
                        XmlElement c = xmlDocument.CreateElement("BLOCKNAME");
                        XmlElement d = xmlDocument.CreateElement("STL_FILENAME");
                        XmlElement e = xmlDocument.CreateElement("SHIFT");
                        XmlElement f = xmlDocument.CreateElement("SAVEPATHC4DFILE");
                        XmlElement g = xmlDocument.CreateElement("SAVEPATHRENDEROUTPUT");

                        XmlText aValue = xmlDocument.CreateTextNode(drawingInfos["customProperty"]);
                        XmlText bValue = xmlDocument.CreateTextNode(drawingInfos["cadName"]);
                        XmlText cValue = xmlDocument.CreateTextNode(item.Item1);
                        XmlText dValue = xmlDocument.CreateTextNode(item.Item2);
                        XmlText eValue = xmlDocument.CreateTextNode(checkBox6.Checked.ToString());
                        XmlText fValue = xmlDocument.CreateTextNode(drawingInfos["cadPath"]);
                        XmlText gValue = xmlDocument.CreateTextNode(drawingInfos["pdfPath"]);

                        //Append nodes to the parent node
                        parentNode.AppendChild(a);
                        parentNode.AppendChild(b);
                        parentNode.AppendChild(c);
                        parentNode.AppendChild(d);
                        parentNode.AppendChild(e);
                        parentNode.AppendChild(f);
                        parentNode.AppendChild(g);

                        a.AppendChild(aValue);
                        b.AppendChild(bValue);
                        c.AppendChild(cValue);
                        d.AppendChild(dValue);
                        e.AppendChild(eValue);
                        f.AppendChild(fValue);
                        g.AppendChild(gValue);

                        xmlElementCounter++;
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("can't create parentNode. (" + item.Item2 + ") Original error: " + ex.Message);
                    }
                }

                try
                {
                    xmlDocument.Save(Path.Combine(Properties.Settings.Default.KEX_STL_PATH, Environment.UserName + ".xml"));
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("xmlDocument.Save FAILED " + "Original error: " + ex.Message);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Can't save the .XML file. " + "Original error: " + ex.Message);
            }
        }


        private ObjectIdCollection GetEntitiesOnLayer(Editor ed, string layerName)
        {
            TypedValue[] tvs = new TypedValue[1] {new TypedValue((int)DxfCode.LayerName, layerName)};
            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            if (psr.Status == PromptStatus.OK)
                return new ObjectIdCollection(psr.Value.GetObjectIds());
            else
                return new ObjectIdCollection();
        }

        private void ExplodeBlocks(Transaction tr, Database db, BlockTableRecord modelspace, ObjectId id, int blockCounter, int level, string originalName, bool explodeNested = true, bool erase = true)
        {
            BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
            ObjectIdCollection toExplode = new ObjectIdCollection();

            try
            {
                //br.ExplodeToOwnerSpace(); //Autocad freeze (hang) with this function if the blockreference have nofields! -> br.HasFields
                DBObjectCollection dboc = new DBObjectCollection();
                br.Explode(dboc);

                if (dboc.Count > 0)
                {
                    foreach (DBObject dbo in dboc)
                    {
                        if (dbo is BlockReference || dbo is Solid3d)
                        {
                            ObjectId oid = modelspace.AppendEntity((Entity)dbo);
                            tr.AddNewlyCreatedDBObject(dbo, true);

                            if (dbo is BlockReference)
                            {
                                toExplode.Add(dbo.ObjectId);
                            }
                            else if (dbo is Solid3d)
                            {
                                try
                                {
                                    //Move Solid3d to a new Layer
                                    Solid3d solid = (Solid3d)tr.GetObject(oid, OpenMode.ForWrite);
                                    try
                                    {
                                        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                                        foreach (ObjectId obj in lt)
                                        {
                                            LayerTableRecord layer = (LayerTableRecord)tr.GetObject(obj, OpenMode.ForRead);
                                            if (layer.Name == solid.Layer)
                                            {
                                                //if the solid layer is frozen oder hidden, delete it
                                                if (!Properties.Settings.Default.KEX_EXPORT_HIDDEN && (layer.IsFrozen || layer.IsHidden || layer.IsOff)) { solid.Erase(); break; }

                                                //Create the new Layer
                                                string newLayerName = Properties.Settings.Default.KEX_GROUP_PREFIX + blockCounter.ToString() + _layerSeperator + originalName + _layerSeperator + solid.Layer;
                                                if (level > 0) newLayerName = Properties.Settings.Default.KEX_GROUP_PREFIX + blockCounter.ToString() + _layerSeperator + originalName + _layerSeperator + level.ToString() + _layerSeperator + _helper.RemoveSpecialCharacters(br.Name.ToUpper()) + _layerSeperator + solid.Layer;
                                                if (!lt.Has(newLayerName))
                                                {
                                                    LayerTableRecord newLayer = new LayerTableRecord();
                                                    newLayer.Name = newLayerName;
                                                    lt.UpgradeOpen();
                                                    lt.Add(newLayer);
                                                    tr.AddNewlyCreatedDBObject(newLayer, true);
                                                }
                                                solid.Layer = newLayerName;
                                                break;
                                            }
                                        }
                                    }
                                    catch (System.Exception acadEx)
                                    {
                                        acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Rename Solid3d-layer failed: " + solid.Layer + " -> Original error: " + acadEx.Message + "\n");
                                    }
                                }
                                catch (System.Exception acadEx)
                                {
                                    acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("New layer name... original error: " + acadEx.Message + "\n");
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception acadEx)
            {
                acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Explode block original error: " + acadEx.Message + "\n");
            }

            if (explodeNested)
            {
                level++;
                foreach (ObjectId bid in toExplode)
                {
                    ExplodeBlocks(tr, db, modelspace, bid, blockCounter, level, originalName, true, erase);
                }
            }

            toExplode.Clear();

            if (erase)
            {
                try
                {
                    br.Erase();
                }
                catch (System.Exception acadEx)
                {
                    acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Erase blockreference failed: " + br.Name + " -> Original error: " + acadEx.Message + "\n" + acadEx.StackTrace + "\n");
                }
            }
        }

        //PROGRESS BAR UPDATE
        private int UpdateProgressBar(double progressStep)
        {
            //Update progressBar
            _progress += progressStep;
            if (_progress > progressBar1.Maximum)
            {
                _progress = progressBar1.Maximum;
            }
            return (int)Math.Round(_progress, 0);
        }

        //CREATE A NEW UNNAMED DRAWING
        private Document CreateNewDrawing(DocumentCollection documentManager)
        {
            try
            {
                string strTemplatePath = "acad.dwt";
                ObjectIdCollection oIdCol = new ObjectIdCollection();

                //Get all not locked layers
                Database db = documentManager.CurrentDocument.Database;
                List<string> notlocked = new List<string>();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    foreach (ObjectId id in layerTable)
                    {
                        var layer = (LayerTableRecord)tr.GetObject(id, OpenMode.ForRead);
                        if (!layer.IsLocked) notlocked.Add(layer.Name);
                    }
                    tr.Commit();
                }

                if (notlocked.Count > 0)
                {
                    string objectfilter = "INSERT";
                    if (Properties.Settings.Default.KEX_SOLID_OUTSIDE_BLOCK)
                    {
                        objectfilter += ",3DSOLID";
                    }
                    TypedValue[] tv = new TypedValue[4] {
                        new TypedValue((int)DxfCode.ViewportVisibility, 0), //0 = Model space
                        new TypedValue((int)DxfCode.Visibility, 0),
                        new TypedValue((int)DxfCode.LayerName, string.Join(",", notlocked)),
                        new TypedValue((int)DxfCode.Start, objectfilter),
                        };
                    SelectionFilter sf = new SelectionFilter(tv);
                    PromptSelectionResult res = documentManager.MdiActiveDocument.Editor.SelectAll(sf);
                    SelectionSet copySset = res.Value;

                    foreach (SelectedObject so in copySset)
                    {
                        oIdCol.Add(so.ObjectId);
                    }

                    copySset = null;
                    res = null;

                    try
                    {
                        Document acNewDoc = documentManager.Add(strTemplatePath);
                        Database acNewDb = acNewDoc.Database;

                        using (DocumentLock acDocLock = acNewDoc.LockDocument())
                        {
                            using (Transaction tr = acNewDb.TransactionManager.StartTransaction())
                            {
                                BlockTable bt = (BlockTable)tr.GetObject(acNewDb.BlockTableId, OpenMode.ForRead);
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(acNewDb.CurrentSpaceId, OpenMode.ForWrite);
                                IdMapping map = new IdMapping();
                                acNewDb.WblockCloneObjects(oIdCol, btr.ObjectId, map, DuplicateRecordCloning.Ignore, false);
                                tr.Commit();
                            }
                        }
                        return acNewDoc;
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Can't copy objects to the new drawing...  Original error: " + ex.Message);
                        return null;
                    }
                }
                else
                {
                    MessageBox.Show("No items found, or everything is on a locked layer");
                    return null;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Move objects to a new drawing failed... Original error:" + ex.Message);
                return null;
            }
        }

        //GET DRAWING INFOS
        private Dictionary<string, string> GetDrawingInfos(Editor ed)
        {
            Dictionary<string, string> drawingInfos = new Dictionary<string, string>();

            //CAD path
            string cadPath = Path.GetDirectoryName(ed.Document.Name);
            try
            {
                string subFolder = Properties.Settings.Default.KEX_C4D_SUB_FOLDER;
                if (subFolder != "")
                {
                    string c4dPath = Path.Combine(cadPath, subFolder);
                    if (!Directory.Exists(c4dPath)) Directory.CreateDirectory(c4dPath);
                    cadPath = c4dPath;
                }
            }
            catch { }
            drawingInfos.Add("cadPath", cadPath);
            drawingInfos.Add("cadName", Path.GetFileNameWithoutExtension(ed.Document.Name));

            //PDF path
            string pdfPath = "";
            bool hasPdfValue = _customPropertiesDict.TryGetValue(Properties.Settings.Default.KEX_RENDERPATH_PROPERTY_NAME, out string pdfValue);
            if (hasPdfValue && Directory.Exists(pdfValue))
            {
                try
                {
                    string subFolder = Properties.Settings.Default.KEX_RENDER_SUB_FOLDER;
                    if (subFolder != "")
                    {
                        pdfValue = Path.Combine(pdfValue, subFolder);
                        if (!Directory.Exists(pdfValue)) Directory.CreateDirectory(pdfValue);
                        pdfPath = pdfValue;
                    }
                }
                catch { }
            }
            drawingInfos.Add("pdfPath", pdfPath);

            //Filename
            string c4dFileName = Path.GetFileNameWithoutExtension(ed.Document.Name).Replace(".", "-");
            drawingInfos.Add("c4dFileName", c4dFileName);

            //Custom property from options
            string customProperty = "";
            bool hasCustomProperty = _customPropertiesDict.TryGetValue(Properties.Settings.Default.KEX_ROOT_PROPERTY_NAME, out string customPropertyValue);
            if (hasCustomProperty) customProperty = customPropertyValue;
            drawingInfos.Add("customProperty", customProperty);

            return drawingInfos;
        }

        //CLEAR RESULTS
        private void ClearResults()
        {
            textBox1.Clear();
            button1.Enabled = true;
            button5.Enabled = false;
            progressBar1.Value = progressBar1.Minimum;
            groupBox4.BackColor = System.Drawing.Color.Transparent;
            groupBox2.Text = "Results";
            label10.Text = "Blocks...";
            label11.Text = "Solids...";
            label9.Text = "Status";
        }

        //OPEN FOLDER BUTTON
        private void Button2_Click(object sender, EventArgs e)
        {
            string openDirectory = Path.Combine(Properties.Settings.Default.KEX_STL_PATH, Environment.UserName);
            bool openDirectoryOk = _helper.OpenDirectory(openDirectory);
            if(!openDirectoryOk)
            {
                MessageBox.Show("Open target directory failed:\n" + openDirectory);
                return;
            }
        }

        //CLEAR RESULTS BUTTON
        private void Button3_Click(object sender, EventArgs e)
        {
            ClearResults();
        }

        //FILEDIA 1 BUTTON
        private void Button4_Click(object sender, EventArgs e)
        {
            acApp.SetSystemVariable("FILEDIA", 1);
            acApp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("FILEDIA varialbe activated (1)\n");
        }

        //FACETRES 0.01 BUTTON
        private void Button7_Click(object sender, EventArgs e)
        {
            SaveFacetresValue(0.01);
        }

        //FACETRES 8 BUTTON
        private void Button8_Click(object sender, EventArgs e)
        {
            SaveFacetresValue(8);
        }

        //FACETRES 10 BUTTON
        private void Button9_Click(object sender, EventArgs e)
        {
            SaveFacetresValue(10);
        }

        //SAVE QUALITY WITH BUTTONS
        private void SaveFacetresValue(double quality)
        {
            Properties.Settings.Default.KEX_QUALITY = quality;
            Properties.Settings.Default.Save();
            textBox4.Text = quality.ToString();
        }

        //CUSTOM PROPERTIES C4D ROOT NAME COMBOBOX
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.KEX_ROOT_PROPERTY_NAME = comboBox1.GetItemText(comboBox1.SelectedItem);
            Properties.Settings.Default.Save();
        }

        //CUSTOM PROPERTIES RENDER PATH COMBOBOX
        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.KEX_RENDERPATH_PROPERTY_NAME = comboBox2.GetItemText(comboBox2.SelectedItem);
            Properties.Settings.Default.Save();
        }

        //SELECT PATH FOR EXPORT BUTTON
        private void Button6_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = Properties.Settings.Default.KEX_STL_PATH;
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    Properties.Settings.Default.KEX_STL_PATH = fbd.SelectedPath;
                    Properties.Settings.Default.Save();
                    textBox3.Text = fbd.SelectedPath;
                }
            }
        }

        //REFRESH PROPERTIES BUTTON
        private void Button10_Click(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            LoadSettings();
        }

        //ABORT EXPORT BUTTON
        private void Button5_Click(object sender, EventArgs e)
        {
            _running = false;
            button1.Enabled = true;
            button5.Enabled = false;
            groupBox4.BackColor = System.Drawing.Color.Transparent;
            label9.Text = "Status";
            label9.Update();
            label9.Refresh();
            progressBar1.Value = progressBar1.Minimum;
            this.Refresh();
        }

        //USE NEW DRAWING CHECKBOX CHANGED
        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            checkBox9.Checked = checkBox8.Checked;
            checkBox9.Enabled = checkBox8.Checked;
        }

        //EXPLODE BLOCKS CHECKBOX CHANGED
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            checkBox2.Checked = checkBox1.Checked;
            //checkBox2.Enabled = checkBox1.Checked;
        }

        //EXCEPTION LAYER CHECKBOX CHANGED
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            textBox2.Enabled = checkBox3.Checked;
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.KEX_USE_TIMEOUT = checkBox10.Checked;
            Properties.Settings.Default.Save();
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.KEX_EXPORT_HIDDEN = checkBox4.Checked;
            Properties.Settings.Default.Save();
        }
    }
}