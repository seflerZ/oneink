/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  OneInk - OneNote Ink Operations COM AddIn
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ClassInterfaceType = System.Runtime.InteropServices.ClassInterfaceType;
using System.Threading;
using System.Threading.Tasks;
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

        static AddIn()
        {
            // Static constructor - this runs BEFORE any instance is created
            // This helps us see if the class is even being loaded
            try
            {
                var logPath = Path.Combine(Path.GetTempPath(), "OneInk.log");
                var msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === STATIC CONSTRUCTOR === Assembly loaded: {System.Reflection.Assembly.GetExecutingAssembly().Location}";
                File.AppendAllText(logPath, msg + Environment.NewLine);
                Debug.WriteLine(msg);
            }
            catch { }
        }

        public AddIn()
        {
            Log("OneInk loaded");

            // For debugging: uncomment to force debugger attach
            // if (!Debugger.IsAttached) { Debugger.Launch(); }
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

                string xml;
                OneNoteApplication.GetPageContent(pageId, out xml, Microsoft.Office.Interop.OneNote.PageInfo.piBinaryData);

                if (string.IsNullOrEmpty(xml))
                {
                    MessageBox.Show(Strings.RetrieveFailed);
                    return;
                }

                var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };
                XDocument doc;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xml), settings))
                    doc = XDocument.Load(reader);
                XNamespace ns = doc.Root.Name.Namespace;

                var inkElements = doc.Descendants(ns + "InkDrawing").ToList();

                if (inkElements.Count == 0)
                {
                    MessageBox.Show(Strings.NoInkStrokes);
                    return;
                }

                foreach (var ink in inkElements)
                {
                    string objectId = ink.Attribute("objectID")?.Value;
                    if (!string.IsNullOrEmpty(objectId))
                        OneNoteApplication.DeletePageContent(pageId, objectId);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(Strings.ErrorClear, ex.Message));
            }
        }

        /// <summary>
        /// Button click handler for deleting ink strokes by selected color.
        /// Parses ISF (Ink Serialized Format) data from each InkDrawing to extract
        /// accurate stroke colors, then deletes matching strokes.
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

                string xml;
                OneNoteApplication.GetPageContent(pageId, out xml, Microsoft.Office.Interop.OneNote.PageInfo.piBinaryData);

                if (string.IsNullOrEmpty(xml))
                {
                    MessageBox.Show(Strings.RetrieveFailed);
                    return;
                }

                var settings = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Ignore };
                XDocument doc;
                using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xml), settings))
                    doc = XDocument.Load(reader);
                XNamespace ns = doc.Root.Name.Namespace;

                var inkElements = doc.Descendants(ns + "InkDrawing").ToList();

                if (inkElements.Count == 0)
                {
                    MessageBox.Show(Strings.NoInkStrokes);
                    return;
                }

                var colorCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var colorObjectIds = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var ink in inkElements)
                {
                    string objectId = ink.Attribute("objectID")?.Value;

                    var dataEl = ink.Element(ns + "Data");
                    string isfBase64 = dataEl?.Value.Trim();
                    string color;

                    if (!string.IsNullOrEmpty(isfBase64))
                    {
                        var isfColors = InkColorExtractor.ExtractInkColors(isfBase64);
                        if (isfColors.Count > 0)
                        {
                            color = isfColors[0];
                        }
                        else
                        {
                            color = ExtractColorFromXml(ink, ns, doc);
                        }
                    }
                    else
                    {
                        color = ExtractColorFromXml(ink, ns, doc);
                    }

                    if (!colorCounts.ContainsKey(color))
                    {
                        colorCounts[color] = 0;
                        colorObjectIds[color] = new List<string>();
                    }
                    colorCounts[color]++;
                    if (!string.IsNullOrEmpty(objectId))
                        colorObjectIds[color].Add(objectId);
                }

                if (colorCounts.Count == 0)
                {
                    MessageBox.Show(Strings.NoInkStrokes);
                    return;
                }

                var items = colorCounts
                    .OrderByDescending(kv => kv.Value)
                    .Select(kv => $"{kv.Key} ({kv.Value} stroke{(kv.Value > 1 ? "s" : "")})")
                    .ToArray();

                using (var dialog = new ColorSelectionDialog(colorCounts.Keys.ToList(), items))
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
