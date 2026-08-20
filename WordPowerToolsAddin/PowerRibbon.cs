using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using OpenXmlPowerTools;

namespace WordPowerToolsAddin
{
    [ComVisible(true)]
    public class PowerRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordPowerToolsAddin.PowerRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #region Ribbon Callbacks
        public void OnMergeClicked(Office.IRibbonControl control)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Word Documents|*.docx", Multiselect = true, Title = "בחר מסמכים למיזוג" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string outputPath = Path.Combine(Path.GetTempPath(), $"Merged_{Guid.NewGuid()}.docx");
                    try
                    {
                        List<Source> sources = new List<Source>();
                        foreach (string file in ofd.FileNames)
                        {
                            sources.Add(new Source(new WmlDocument(file), true));
                        }
                        DocumentBuilder.BuildDocument(sources, outputPath);
                        OpenInWord(outputPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"שגיאה במיזוג: {ex.Message}");
                    }
                }
            }
        }

        public void OnReplaceClicked(Office.IRibbonControl control)
        {
            string search = "טקסט ישן";
            string replace = "טקסט חדש";
            ProcessActiveDocument((wmlDoc) => TextReplacer.SearchAndReplace(wmlDoc, search, replace, true));
        }

        public void OnAcceptRevisionsClicked(Office.IRibbonControl control)
        {
            ProcessActiveDocument((wmlDoc) => RevisionAccepter.AcceptRevisions(wmlDoc));
        }

        public void OnSimplifyClicked(Office.IRibbonControl control)
        {
            ProcessActiveDocument((wmlDoc) =>
            {
                SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                {
                    RemoveComments = true,
                    RemoveContentControls = true,
                    RemoveBookmarks = true,
                    RemoveWebHidden = true
                };
                return MarkupSimplifier.SimplifyMarkup(wmlDoc, settings);
            });
        }

        public void OnExtractTextClicked(Office.IRibbonControl control)
        {
            string tempDoc = SaveActiveDocumentToTemp();
            if (string.IsNullOrEmpty(tempDoc)) return;

            try
            {
                WmlDocument wmlDoc = new WmlDocument(tempDoc);
                string plainText = WmlToTextConverter.Convert(wmlDoc, true);
                string txtPath = Path.Combine(Path.GetTempPath(), $"ExtractedText_{Guid.NewGuid()}.txt");
                File.WriteAllText(txtPath, plainText);
                System.Diagnostics.Process.Start(txtPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"שגיאה בחילוץ טקסט: {ex.Message}");
            }
        }
        #endregion

        #region Helpers
        private string SaveActiveDocumentToTemp()
        {
            Word.Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            if (activeDoc == null)
            {
                MessageBox.Show("אין מסמך פעיל לעבוד עליו.");
                return null;
            }
            string tempPath = Path.Combine(Path.GetTempPath(), $"Copy_{Guid.NewGuid()}.docx");
            activeDoc.SaveAs2(tempPath);
            return tempPath;
        }

        private void ProcessActiveDocument(Func<WmlDocument, WmlDocument> powerToolAction)
        {
            string tempFile = SaveActiveDocumentToTemp();
            if (string.IsNullOrEmpty(tempFile)) return;

            try
            {
                WmlDocument originalDoc = new WmlDocument(tempFile);
                WmlDocument processedDoc = powerToolAction(originalDoc);
                string resultPath = Path.Combine(Path.GetTempPath(), $"Result_{Guid.NewGuid()}.docx");
                processedDoc.SaveAs(resultPath);
                OpenInWord(resultPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"אירעה שגיאה בעיבוד: {ex.Message}");
            }
        }

        private void OpenInWord(string filePath)
        {
            if (File.Exists(filePath))
            {
                Globals.ThisAddIn.Application.Documents.Open(filePath);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}