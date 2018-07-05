// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RecordDocumentManager.cs" company="Montrium">
//   MIT Licence
// </copyright>
// <summary>
//   
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections.Generic;
    using Microsoft.SharePoint;


    /// <summary>
    /// Manages the records
    /// </summary>
    public class RecordDocumentManager : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="RecordDocumentManager"/> class.
        /// </summary>
        /// <param name="dropOffLibraryRecords">The drop off library records.</param>
        /// <param name="enabledLibraries">The SP web lists.</param>
        /// <param name="routingRules">The routing rules.</param>
        public RecordDocumentManager(
            SPListItemCollection dropOffLibraryRecords, List<SPList> enabledLibraries, SPListItemCollection routingRules)
        {
            if (dropOffLibraryRecords == null) throw new ArgumentNullException("dropOffLibraryRecords");
            if (enabledLibraries == null) throw new ArgumentException("enabledLibraries");
            if (routingRules == null) throw new ArgumentNullException("routingRules");

            List<RecordDocument> rcdList = new List<RecordDocument>();
            foreach (SPListItem DropOffLibraryRecord in dropOffLibraryRecords)
                rcdList.Add(new RecordDocument(enabledLibraries, DropOffLibraryRecord, routingRules));

            this.RecordDocuments = rcdList;
        }
        #endregion

        #region DestructorDisposable
        /// <summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="RecordDocumentManager"/> is reclaimed by garbage collection.
        /// </summary>
        ~RecordDocumentManager()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets the record documents.</summary>
        public List<RecordDocument> RecordDocuments { get; internal set; }
        #endregion

        #region PublicMethods
        /// <summary>
        /// Scans the fields.
        /// </summary>
        /// <param name="dropOffLibrary">The drop off library.</param>
        /// <param name="dropOffLibraryUrl">The drop off library URL.</param>
        public void ScanFields(SPList dropOffLibrary, string dropOffLibraryUrl)
        {
            // rd.SpListItem, rlm.DropOffLibrary, rlm.DropOffLibUrl, rd.XmlProperties
            foreach (RecordDocument rd in this.RecordDocuments)
            {
                RecordFieldManager rfm = new RecordFieldManager(
                    rd.ListItem, dropOffLibrary, dropOffLibraryUrl, rd.XmlProperties);
                rfm.PrintToLogFile(rd.LogFile);
            }
        }

        /// <summary>
        /// Moves the file to library.
        /// </summary>
        public void MoveFileToLibrary()
        {
            foreach (RecordDocument recordDocument in this.RecordDocuments)
            {
                string newUrl;
                SPListItem routingRule;
                SPList newLib;

                if (recordDocument.HasLibrary)
                {
                    // child level library
                    routingRule = recordDocument.RoutingRule;
                    newUrl = routingRule.Web.Url + "/";
                    newLib = recordDocument.CandidateLibrary;

                    if (recordDocument.HasRoutingRule)
                    {
                        // child level library with rule
                        if (routingRule["Target Folder"] == null)
                        {
                            // child level library with rule without folder
                            newUrl += routingRule["Target Library"] + "/";
                            recordDocument.MoveToLibraryWithoutFolder(newUrl);
                        }
                        else
                        {
                            // child level library with rule with folder
                            SPFolder newSpFolder = newLib.Folders[0].Folder;
                            newUrl = newSpFolder.Url;
                            recordDocument.MoveToLibraryWithFolder(newSpFolder);
                        }
                    }
                    else
                    {
                        // parent level library only - without rule
                        newUrl = recordDocument.CandidateLibrary.ParentWebUrl + "/"
                                 + recordDocument.CandidateLibrary.Title;
                    }
                }
                else if (recordDocument.HasParentLibrary)
                {
                    // parent level library
                    routingRule = recordDocument.ParentRoutingRule;
                    newUrl = routingRule.Web.Url + "/";
                    newLib = recordDocument.ParentCandidateLibrary;

                    if (recordDocument.HasParentRoutingRule)
                    {
                        // parent level library with rule
                        if (routingRule["Target Folder"] == null)
                        {
                            // parent level library with rule without folder
                            newUrl += routingRule["Target Library"] + "/";
                            recordDocument.MoveToLibraryWithoutFolder(newUrl);
                        }
                        else
                        {
                            // parent level library with rule with folder
                            SPFolder newSpFolder = newLib.Folders[0].Folder;
                            newUrl = newSpFolder.Url;
                            recordDocument.MoveToLibraryWithFolder(newSpFolder);
                        }
                    }
                    else
                    {
                        // parent level library only - without rule
                        newUrl = recordDocument.ParentCandidateLibrary.ParentWebUrl + "/" + recordDocument.ParentCandidateLibrary.Title;
                    }
                }
                else
                {
                    // unknown level library
                    continue;
                }

                // if (newLib != null) newLib.Update();
            }
        }
        #endregion

        #region PublicOverride
        /// <summary>
        /// Returns a <see cref="System.String"/> that represents this instance.
        /// </summary>
        /// <returns>
        /// A <see cref="System.String"/> that represents this instance.
        /// </returns>
        public override string ToString()
        {
            string s = "FileCount:\t\t" + this.RecordDocuments.Count + "\n\n";
            foreach (RecordDocument rdc in this.RecordDocuments)
                s += rdc.ToString();
            return s;
        }
        #endregion

        #region PublicDisposable
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region PrivateDisposable
        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        /// <param name="isDisposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        private void Dispose(bool isDisposing)
        {
            // Check if Dispose has been called
            if (!this.isDisposed)
            {
                // dispose managed and unmanaged resources
                if (isDisposing)
                {
                    // managed resources clean
                    this.RecordDocuments = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion
    }
}
