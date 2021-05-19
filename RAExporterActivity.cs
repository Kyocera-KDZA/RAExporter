using Laserfiche.Custom.Activities;
using Laserfiche.Custom.Activities.Design;
using Laserfiche.RepositoryAccess;
using Laserfiche.Workflow.Activities;
using Laserfiche.Workflow.ComponentModel;
using LFSO104Lib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Workflow.ComponentModel;
using Laserfiche.DocumentServices;
using Laserfiche.Connection;

namespace RAExporter
{
    public class RAExporterActivity : CustomSingleEntryActivity
    {
        public const string ExtensionFileName = "ExtensionFile";

        /// <summary>The name of the Entry Path property.</summary>
        public const string EntryPathName = "EntryPath";

        private string pathInput = "";
        private string extensionInput = "";

        public string PathInput
        {
            get { return this.pathInput; }
            set { this.pathInput = value; }
        }

        public string ExtensionInput
        {
            get { return this.extensionInput; }
            set { this.extensionInput = value; }
        }

        /// <summary>
        /// Called when the activity is run by the workflow server. Implement the logic of your activity in this method. 
        /// Access methods for setting tokens, getting token values, and other functions from the base class or the execution 
        /// context parameter. 
        /// </summary>
        protected override ActivityExecutionStatus Execute(ActivityExecutionContext executionContext)
        {
            using (ConnectionWrapper wrapper = executionContext.OpenConnectionRA92())
            {
                Session session = (Session)wrapper.Connection;
                // Note: You must add the Laserfiche.RepositoryAccess reference to this project for this to work.
                try
                {
                    ILFDatabase database = (ILFDatabase)wrapper.Database;

                    LaserficheEntryInfo entryInfo = this.GetEntryInformation(executionContext);
                    ILFEntry entry = (ILFEntry)database.GetEntryByID(entryInfo.Id);
                    //entry.Name = this.PathInput;

                    //database.GetAndLockEntryByID(uInput);
                    ILFDocument document = (ILFDocument)entry;

                    DocumentInfo docInfo = Document.GetDocumentInfo(entry.ID, session);

                    string expoExtension = this.ResolveTokensInText(executionContext, this.ExtensionInput);
                    string expoPath = this.ResolveTokensInText(executionContext, this.PathInput);

                    if (expoPath != "")
                    {
                        expoPath += docInfo.Name;

                        if (expoExtension == "tiff")
                        {

                            try
                            {
                                //DocumentInfo mydoc = Document.GetDocumentInfo(GetEntryInformation(executionContext).Id, session);
                                DocumentExporter exporter = new DocumentExporter();
                                exporter.PageFormat = DocumentPageFormat.Tiff;
                                exporter.ExportElecDoc(docInfo, (expoPath + ".tiff"));
                            }
                            catch (Exception ex)
                            {
                                this.TrackError(ex);
                            }

                        }
                        else if (expoExtension == "pdf")
                        {
                            try
                            {
                                DocumentExporter exporter = new DocumentExporter();
                                PageSet docPgs = docInfo.AllPages;
                                exporter.ExportPdf(docInfo, docPgs, PdfExportOptions.IncludeText, (expoPath + ".pdf"));
                            }
                            catch (Exception ex)
                            {
                                this.TrackError(ex);
                            }

                        }
                    }

                }
                catch (Exception ex)
                {
                    this.TrackError(ex);
                    return ActivityExecutionStatus.Closed;
                }
            }

            return base.Execute(executionContext);
        }


    }
}
