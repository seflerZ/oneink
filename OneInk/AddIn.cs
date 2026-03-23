/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  OneInk - OneNote Ink Operations COM AddIn
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ClassInterfaceType = System.Runtime.InteropServices.ClassInterfaceType;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using Application = Microsoft.Office.Interop.OneNote.Application;
using IRibbonControl = Microsoft.Office.Core.IRibbonControl;
using IStream = System.Runtime.InteropServices.ComTypes.IStream;

#pragma warning disable CS3003 // Type is not CLS-compliant

namespace OneInk
{
    /// <summary>
    /// Main COM AddIn class for OneNote ink operations.
    /// Implements IDTExtensibility2 for COM add-in lifecycle and IRibbonExtensibility for ribbon UI.
    /// </summary>
    [ComVisible(true)]
    [Guid("E1F2A3B4-CF2D-409B-B65A-BDBACB9F21DC"), ProgId("OneInk.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "OneInk.log");

        private static void Log(string message)
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logLine = $"[{timestamp}] {message}";
                File.AppendAllText(LogPath, logLine + Environment.NewLine);
                Debug.WriteLine(logLine);
            }
            catch { }
        }

        /// <summary>
        /// Reference to the OneNote application object.
        /// </summary>
        protected Application OneNoteApplication { get; set; }

        private static IRibbonUI _ribbon;
        private static int _selectedDensityIndex = 1; // default to medium

        public AddIn()
        {
            Log("OneInk loaded");
        }

        /// <summary>
        /// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon.
        /// </summary>
        /// <param name="RibbonID">The ribbon ID.</param>
        /// <returns>XML string defining the ribbon UI.</returns>
        public string GetCustomUI(string RibbonID)
        {
            try
            {
                return Properties.Resources.ribbon;
            }
            catch (Exception ex)
            {
                Log($"GetCustomUI ERROR: {ex}");
                throw;
            }
        }

        /// <summary>
        /// Called by OneNote when the ribbon is loaded. Stores reference to ribbon UI.
        /// </summary>
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Log("Ribbon loaded");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>
        /// Cleanup when add-in is shutting down.
        /// </summary>
        /// <param name="custom">Custom parameters.</param>
        public void OnBeginShutdown(ref Array custom)
        {
        }

        /// <summary>
        /// Called upon startup. Keeps a reference to the current OneNote application object.
        /// </summary>
        /// <param name="Application">The application object.</param>
        /// <param name="ConnectMode">The connection mode.</param>
        /// <param name="AddInInst">The add-in instance.</param>
        /// <param name="custom">Custom parameters.</param>
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                SetOneNoteApplication((Application)Application);
            }
            catch (Exception ex)
            {
                Log($"OnConnection ERROR: {ex}");
                throw;
            }
        }

        public void SetOneNoteApplication(Application application)
        {
            OneNoteApplication = application;
        }

        /// <summary>
        /// Cleanup when add-in is disconnected.
        /// </summary>
        /// <param name="RemoveMode">The disconnection mode.</param>
        /// <param name="custom">Custom parameters.</param>
        [SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                OneNoteApplication = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>
        /// Button click handler for clearing all ink on current page.
        /// If the user has ink selected on the page, only selected ink is cleared.
        /// Otherwise, all ink on the page is cleared.
        /// </summary>
        /// <param name="control">The ribbon control that triggered the action.</param>
        public void ClearInkButtonClicked(IRibbonControl control)
        {
            try
            {
                if (OneNoteApplication == null)
                {
                    MessageBox.Show(Strings.AppNotAvailable);
                    return;
                }

                string pageId = OneNoteApplication.Windows.CurrentWindow.CurrentPageId;

                // Step 1: Get selection info using piSelection (fast ~20ms)
                OneNoteApplication.GetPageContent(pageId, out string xmlSelection, Microsoft.Office.Interop.OneNote.PageInfo.piSelection);

                var selSettings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };
                XDocument docSel;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlSelection ?? ""), selSettings))
                    docSel = XDocument.Load(reader);
                XNamespace ns = docSel.Root.Name.Namespace;

                // Get selected ink object IDs (marked with selected="all")
                var selectedObjectIds = new HashSet<string>(
                    docSel.Descendants(ns + "InkDrawing")
                          .Where(e => e.Attribute("selected")?.Value == "all")
                          .Select(e => e.Attribute("objectID")?.Value ?? "")
                          .Where(id => !string.IsNullOrEmpty(id))
                );
                bool hasSelection = selectedObjectIds.Count > 0;

                // Step 2: Get all ink object IDs using piBasic (fast ~20ms, no binary data)
                OneNoteApplication.GetPageContent(pageId, out string xmlBasic, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);

                XDocument docBasic;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlBasic ?? ""), selSettings))
                    docBasic = XDocument.Load(reader);

                var allInkIds = docBasic.Descendants(ns + "InkDrawing")
                                   .Select(e => e.Attribute("objectID")?.Value ?? "")
                                   .Where(id => !string.IsNullOrEmpty(id))
                                   .ToList();

                // Step 3: Delete ink by objectId (DeletePageContent is fast)
                int deletedCount = 0;
                foreach (var objectId in allInkIds)
                {
                    if (hasSelection && !selectedObjectIds.Contains(objectId))
                        continue;

                    OneNoteApplication.DeletePageContent(pageId, objectId);
                    deletedCount++;
                }

                if (deletedCount == 0)
                    MessageBox.Show(hasSelection ? Strings.NoInkStrokesInSelection : Strings.NoInkStrokes);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(Strings.ErrorClear, ex.Message));
            }
        }

        /// <summary>
        /// Button click handler for converting selected ink to dashed/dotted lines.
        /// Breaks each stroke into segments with gaps at regular intervals.
        /// If the user has ink selected, only selected ink is converted.
        /// </summary>
        /// <param name="control">The ribbon control that triggered the action.</param>
        public void ToDashedInkButtonClicked(IRibbonControl control)
        {
            ExecuteToDashed();
        }

        private void ExecuteToDashed()
        {
            try
            {
                if (OneNoteApplication == null)
                {
                    MessageBox.Show(Strings.AppNotAvailable);
                    return;
                }

                string pageId = OneNoteApplication.Windows.CurrentWindow.CurrentPageId;

                // Get spacing from ribbon dropdown selection
                float dashFraction;
                switch (_selectedDensityIndex)
                {
                    case 0: dashFraction = 0.025f; break;
                    case 2: dashFraction = 0.10f; break;
                    default: dashFraction = 0.05f; break;
                }

                // Step 1: Get selection info using piSelection (fast ~20ms)
                OneNoteApplication.GetPageContent(pageId, out string xmlSelection, Microsoft.Office.Interop.OneNote.PageInfo.piSelection);

                var selSettings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };
                XDocument docSel;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlSelection ?? ""), selSettings))
                    docSel = XDocument.Load(reader);
                XNamespace ns = docSel.Root.Name.Namespace;

                // Get selected ink object IDs (marked with selected="all")
                var selectedObjectIds = new HashSet<string>(
                    docSel.Descendants(ns + "InkDrawing")
                          .Where(e => e.Attribute("selected")?.Value == "all")
                          .Select(e => e.Attribute("objectID")?.Value ?? "")
                          .Where(id => !string.IsNullOrEmpty(id))
                );

                if (selectedObjectIds.Count == 0)
                {
                    MessageBox.Show("请先选择要转换的墨迹（套索工具框选）");
                    return;
                }

                // Step 2: Get page structure using piBasic (fast ~20ms)
                OneNoteApplication.GetPageContent(pageId, out string xmlBasic, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);

                XDocument docBasic;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlBasic ?? ""), selSettings))
                    docBasic = XDocument.Load(reader);

                // Build a dictionary of selected ink elements (for Position, Size, lastModifiedTime)
                var selectedInkElements = docBasic.Descendants(ns + "InkDrawing")
                    .Where(e => selectedObjectIds.Contains(e.Attribute("objectID")?.Value ?? ""))
                    .ToList();

                // Step 3: For each selected ink, get binary data and convert
                int convertedCount = 0;
                var inkXmlParts = new List<XElement>();

                foreach (var inkEl in selectedInkElements)
                {
                    string objectId = inkEl.Attribute("objectID")?.Value;
                    string lastModified = inkEl.Attribute("lastModifiedTime")?.Value ?? "";

                    if (string.IsNullOrEmpty(objectId))
                        continue;

                    // Get binary data for this specific ink using GetBinaryPageContent
                    OneNoteApplication.GetBinaryPageContent(pageId, objectId, out string isfBase64);
                    if (string.IsNullOrEmpty(isfBase64))
                        continue;

                    // Convert to dashed
                    string dashedBase64 = InkDashedConverter.ConvertToDashed(isfBase64, dashFraction, dashFraction);
                    if (string.IsNullOrEmpty(dashedBase64))
                        continue;

                    // Build ink XML element with Position, Size, and new Data
                    var inkXmlEl = new XElement(ns + "InkDrawing",
                        new XAttribute("objectID", objectId),
                        new XAttribute("lastModifiedTime", lastModified),
                        inkEl.Element(ns + "Position"),
                        inkEl.Element(ns + "Size"),
                        new XElement(ns + "Data", dashedBase64)
                    );
                    inkXmlParts.Add(inkXmlEl);
                    convertedCount++;
                }

                if (convertedCount == 0)
                {
                    MessageBox.Show(Strings.NoInkStrokesInSelection);
                    return;
                }

                // Step 4: Update page with partial XML (just the modified ink elements)
                var pageEl = new XElement(ns + "Page",
                    new XAttribute("ID", pageId)
                );
                foreach (var inkEl in inkXmlParts)
                    pageEl.Add(inkEl);

                var pageDoc = new XDocument(pageEl);
                string pageXml = pageDoc.ToString(SaveOptions.DisableFormatting);
                OneNoteApplication.UpdatePageContent(pageXml);
            }
            catch (Exception ex)
            {
                Log($"ExecuteToDashed ERROR: {ex}");
                MessageBox.Show(string.Format(Strings.ErrorDashed, ex.Message));
            }
        }

        /// <summary>
        /// Button click handler for deleting ink strokes by selected color.
        /// Parses ISF (Ink Serialized Format) data from each InkDrawing to extract
        /// accurate stroke colors, then deletes matching strokes.
        /// If the user has ink selected, only selected ink is considered.
        /// </summary>
        /// <param name="control">The ribbon control that triggered the action.</param>
        public void SelectInkColorButtonClicked(IRibbonControl control)
        {
            try
            {
                if (OneNoteApplication == null)
                {
                    MessageBox.Show(Strings.AppNotAvailable);
                    return;
                }

                string pageId = OneNoteApplication.Windows.CurrentWindow.CurrentPageId;

                var selSettings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };

                // Step 1: Get selection info using piSelection (fast ~20ms)
                OneNoteApplication.GetPageContent(pageId, out string xmlSelection, Microsoft.Office.Interop.OneNote.PageInfo.piSelection);

                XDocument docSel;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlSelection ?? ""), selSettings))
                    docSel = XDocument.Load(reader);
                XNamespace ns = docSel.Root.Name.Namespace;

                var selectedObjectIds = new HashSet<string>(
                    docSel.Descendants(ns + "InkDrawing")
                          .Where(e => e.Attribute("selected")?.Value == "all")
                          .Select(e => e.Attribute("objectID")?.Value ?? "")
                          .Where(id => !string.IsNullOrEmpty(id))
                );
                bool hasSelection = selectedObjectIds.Count > 0;

                // Step 2: Get page structure using piBasic (fast ~20ms)
                OneNoteApplication.GetPageContent(pageId, out string xmlBasic, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);

                XDocument docBasic;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xmlBasic ?? ""), selSettings))
                    docBasic = XDocument.Load(reader);

                var allInkElements = docBasic.Descendants(ns + "InkDrawing").ToList();
                if (allInkElements.Count == 0)
                {
                    MessageBox.Show(Strings.NoInkStrokes);
                    return;
                }

                // Determine which inks to check
                var inksToCheck = hasSelection
                    ? allInkElements.Where(e => selectedObjectIds.Contains(e.Attribute("objectID")?.Value ?? "")).ToList()
                    : allInkElements;

                // Step 3: Get binary data for each ink in parallel
                var objectIds = inksToCheck
                    .Select(e => e.Attribute("objectID")?.Value ?? "")
                    .Where(id => !string.IsNullOrEmpty(id))
                    .ToList();

                var parallelResults = new System.Collections.Concurrent.ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                System.Threading.Tasks.Parallel.ForEach(objectIds,
                    new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = 8 },
                    objectId =>
                {
                    var app = (Application)Activator.CreateInstance(System.Type.GetTypeFromProgID("OneNote.Application"));
                    app.GetBinaryPageContent(pageId, objectId, out string isfBase64);
                    if (!string.IsNullOrEmpty(isfBase64))
                    {
                        var isfColors = InkColorExtractor.ExtractInkColors(isfBase64);
                        string color = isfColors.Count > 0 ? isfColors[0] : "#000000";
                        parallelResults[objectId] = color;
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                });

                var colorCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var colorObjectIds = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var kvp in parallelResults)
                {
                    string color = kvp.Value;
                    if (!colorCounts.ContainsKey(color))
                    {
                        colorCounts[color] = 0;
                        colorObjectIds[color] = new List<string>();
                    }
                    colorCounts[color]++;
                    colorObjectIds[color].Add(kvp.Key);
                }

                if (colorCounts.Count == 0)
                {
                    MessageBox.Show(hasSelection ? Strings.NoInkStrokesInSelection : Strings.NoInkStrokes);
                    return;
                }

                using (var dialog = new ColorSelectionDialog(colorCounts.Keys.OrderByDescending(c => c).ToList()))
                {
                    try
                    {
                        var window = OneNoteApplication.Windows.CurrentWindow;
                        var hwnd = new IntPtr(Convert.ToInt64(window.WindowHandle));
                        dialog.ShowDialog(new OneNoteWindowOwner(hwnd));
                    }
                    catch
                    {
                        dialog.ShowDialog();
                    }

                    if (dialog.DialogResult != DialogResult.OK || dialog.SelectedColor == null)
                        return;

                    string selectedColor = dialog.SelectedColor;
                    var idsToDelete = colorObjectIds[selectedColor];

                    foreach (string objectId in idsToDelete)
                        OneNoteApplication.DeletePageContent(pageId, objectId);
                }
            }
            catch (Exception ex)
            {
                Log($"SelectInkColorButtonClicked ERROR: {ex}");
                MessageBox.Show(string.Format(Strings.ErrorDelete, ex.Message));
            }
        }

        /// <summary>
        /// Fallback color extraction from XML attributes when ISF parsing is unavailable.
        /// </summary>
        private static string ExtractColorFromXml(XElement ink, XNamespace ns, XDocument doc)
        {
            // Try quickStyleIndex mapping via QuickStyleDef
            var quickStyleMap = new Dictionary<int, string>(32);
            foreach (var def in doc.Descendants(ns + "QuickStyleDef"))
            {
                var indexAttr = def.Attribute("index");
                var colorAttr = def.Attribute("color");
                if (indexAttr != null && colorAttr != null && int.TryParse(indexAttr.Value, out int idx))
                    quickStyleMap[idx] = colorAttr.Value;
            }

            var qsiAttr = ink.Attribute("quickStyleIndex");
            if (qsiAttr != null && int.TryParse(qsiAttr.Value, out int qsi) && quickStyleMap.TryGetValue(qsi, out string qsColor))
                return qsColor;

            // Try Stroke element color attribute
            var strokeEl = ink.Element(ns + "Stroke");
            if (strokeEl != null)
            {
                var strokeColor = strokeEl.Attribute("color")?.Value;
                if (!string.IsNullOrEmpty(strokeColor))
                    return strokeColor;
            }

            return "#000000";
        }

        /// <summary>
        /// Returns the localized label for a ribbon control.
        /// </summary>
        public string GetLabel(IRibbonControl control)
        {
            try
            {
                string label;
                switch (control.Id)
                {
                    case "tabOneInk": label = Strings.RibbonTabLabel; break;
                    case "groupInkTools": label = Strings.RibbonGroupLabel; break;
                    case "buttonClearInk": label = Strings.ButtonClearInkLabel; break;
                    case "buttonDeleteByColor": label = Strings.ButtonDeleteByColorLabel; break;
                    case "buttonToDashed": label = GetCurrentDashedLabel(); break;
                    default: label = null; break;
                }
                return label;
            }
            catch (Exception ex)
            {
                Log($"GetLabel({control.Id}) ERROR: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Returns the localized screentip for a ribbon control.
        /// </summary>
        public string GetScreentip(IRibbonControl control)
        {
            try
            {
                string tip;
                switch (control.Id)
                {
                    case "buttonClearInk": tip = Strings.ButtonClearInkScreentip; break;
                    case "buttonDeleteByColor": tip = Strings.ButtonDeleteByColorScreentip; break;
                    case "buttonToDashed": tip = Strings.ButtonToDashedScreentip; break;
                    default: tip = null; break;
                }
                return tip;
            }
            catch (Exception ex)
            {
                Log($"GetScreentip({control.Id}) ERROR: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Called when Dense menu item is clicked.
        /// </summary>
        public void OnMenuDenseClicked(IRibbonControl control)
        {
            _selectedDensityIndex = 0;
            Log($"OnMenuDenseClicked");
            if (_ribbon != null)
                _ribbon.InvalidateControl("buttonToDashed");
        }

        /// <summary>
        /// Called when Medium menu item is clicked.
        /// </summary>
        public void OnMenuMediumClicked(IRibbonControl control)
        {
            _selectedDensityIndex = 1;
            Log($"OnMenuMediumClicked");
            if (_ribbon != null)
                _ribbon.InvalidateControl("buttonToDashed");
        }

        /// <summary>
        /// Called when Sparse menu item is clicked.
        /// </summary>
        public void OnMenuSparseClicked(IRibbonControl control)
        {
            _selectedDensityIndex = 2;
            Log($"OnMenuSparseClicked");
            if (_ribbon != null)
                _ribbon.InvalidateControl("buttonToDashed");
        }

        private static string GetCurrentDashedLabel()
        {
            string densityName;
            switch (_selectedDensityIndex)
            {
                case 0: densityName = Strings.DashedDense; break;
                case 2: densityName = Strings.DashedSparse; break;
                default: densityName = Strings.DashedMedium; break;
            }
            return $"{Strings.ButtonToDashedLabel}（{densityName}）";
        }

        /// <summary>
        /// Returns the appropriate image for the ToDashed button based on selected density.
        /// </summary>
        public IStream GetToDashedImage(IRibbonControl control)
        {
            string densityImage;
            switch (_selectedDensityIndex)
            {
                case 0: densityImage = "ToDashedDense.png"; break;
                case 2: densityImage = "ToDashedSparse.png"; break;
                default: densityImage = "ToDashedMedium.png"; break;
            }
            return GetImage(densityImage);
        }

        /// <summary>
        /// Specified as the loadImage callback in Ribbon.xml, this method returns the
        /// image to display on the ribbon button.
        /// </summary>
        /// <param name="imageName">The name of the image to retrieve.</param>
        /// <returns>IStream containing the PNG image data.</returns>
        public IStream GetImage(string imageName)
        {
            try
            {
                var assemblyDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                var imagePath = Path.Combine(assemblyDir, "Resources", imageName);
                if (!File.Exists(imagePath))
                    imagePath = Path.Combine(assemblyDir, imageName);
                if (!File.Exists(imagePath))
                {
                    return null;
                }
                using (var bitmap = new Bitmap(imagePath))
                    return bitmap.GetReadOnlyStream();
            }
            catch
            {
                return null;
            }
        }
    }

    internal class OneNoteWindowOwner : IWin32Window
    {
        public OneNoteWindowOwner(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; }
    }
}
