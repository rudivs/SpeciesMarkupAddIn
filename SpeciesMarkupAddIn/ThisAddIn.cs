using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SpeciesMarkupAddIn
{
    public partial class ThisAddIn
    {
        private bool markupEnabled;
        private Microsoft.Office.Tools.CustomTaskPane taxonMarkupPanel;
        public TaxonPanel myTaxonPanel;
        public string currentText;
        public TaxonList currentBatch;
        public Taxon currentTaxon;
        Serializer serializer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            this.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);

            serializer = new Serializer();
            if (!Deserialize())
            {
                currentBatch = new TaxonList();
            }
            if (currentBatch.Count > 0)
            {
                currentTaxon = currentBatch.Current;
            }
            else
            {
                currentTaxon = new Taxon();
                currentBatch.Add(currentTaxon);
            }
            
        }

        void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            Serialize();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        internal void setDisplayFlag(bool value)
        {
            markupEnabled = value;
            if (markupEnabled)
            {
                taxonMarkupPanel.Visible = true;
            }
            else
            {
                taxonMarkupPanel.Visible = false;
            }
        }

        void Application_DocumentChange()
        {
            LoadTaxonPanel();
        }

        internal void LoadTaxonPanel()
        {
            if (taxonMarkupPanel != null)
            {
                this.CustomTaskPanes.Remove(taxonMarkupPanel);
            }

            myTaxonPanel = new TaxonPanel();
            taxonMarkupPanel = this.CustomTaskPanes.Add(myTaxonPanel, "Taxon Markup Panel");
            taxonMarkupPanel.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            taxonMarkupPanel.Width = 370;
            this.myTaxonPanel.LoadCurrentTaxon();
        }

        void Application_WindowSelectionChange(Word.Selection Sel)
        {
            if (markupEnabled)
            {
                currentText = Sel.Text;
            }
        }

        internal bool UpdateTrackingNumber()
        {
            try
            {
                this.currentTaxon.TrackingNumber = Path.GetFileNameWithoutExtension(this.Application.ActiveDocument.Name);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception)
            {
                throw;
            }
        }

        internal bool Serialize()
        {
            try
            {
                string filename = CollectionData.BatchFile;
                serializer.SerializeObject(filename, currentBatch);
                return true;
            }
            catch (DirectoryNotFoundException)
            {
                // MessageBox.Show("Serialization error (directory not found). " + ex,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return false;
            }
            catch (IOException)
            {
                // MessageBox.Show("Serialization error (error accessing file). " + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (Exception)
            {
                throw;
            }
        }

        internal bool Deserialize()
        {
            try
            {
                string filename = CollectionData.BatchFile;
                currentBatch = serializer.DeSerializeObject(filename);
                return true;
            }
            catch (DirectoryNotFoundException)
            {
                // MessageBox.Show("Deserialization error (directory not found). " + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (FileNotFoundException)
            {
                // MessageBox.Show("Deserialization error (file not found). " + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (IOException)
            {
                //MessageBox.Show("Deserialization error (error accessing file). " + ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show("Error loading batch file. This is most likely due to invalid characters in one or more fields." + Environment.NewLine + ex.ToString(), "Error", MessageBoxButtons.OK,MessageBoxIcon.Error);
                return false;
            }
            catch (Exception)
            {
                throw;
            }
        }

        internal void ResetCurrentBatch(string confirmationMessage)
        {
            DialogResult result = DialogResult.No;
            if (confirmationMessage == CollectionData.DeleteAdHoc)
            {
                result = MessageBox.Show(confirmationMessage,"Clear batch",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2);
            }
            if (confirmationMessage == CollectionData.DeleteAfterExport)
            {
                result = MessageBox.Show(confirmationMessage, "Export complete", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            }

            if (result == DialogResult.Yes)
            {
                if (File.Exists(CollectionData.BatchFile))
                {
                    File.Delete(CollectionData.BatchFile);
                }
                this.currentBatch.Reset();
                currentTaxon = new Taxon();
                currentBatch.Add(currentTaxon);
                this.myTaxonPanel.LoadCurrentTaxon();
            }
        }
    }
}
