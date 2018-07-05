// -----------------------------------------------------------------------
// <copyright file="RecordCentreManager.cs" company="Montrium">
// MIT License
// </copyright>
// -----------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections.Generic;
    using Microsoft.SharePoint;

    /// <summary>
    /// Manages The Record Center
    /// </summary>
    public class RecordCentreManager : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="RecordCentreManager"/> class.
        /// </summary>
        /// <param name="url">The URL.</param>
        public RecordCentreManager(string url)
        {
            if (String.IsNullOrEmpty(url)) throw new ArgumentNullException("url");

            using (RecordLibraryManager rlm = new RecordLibraryManager(url))
            using (RoutingRulesManager rrm = new RoutingRulesManager(url))
            {
                rlm.PrintToLogFile(@"C:\Users\eescobar\Desktop\.txt");
                rrm.PrintToLogFile(@"C:\Users\eescobar\Desktop\.txt");

                this.EnabledLibraries = rlm.EnabledLibraries;
                this.DropOffLibrary = rlm.DropOffLibrary;
                this.DropOffLibraryDocuments = rlm.DropOffLibrary.GetItems();

                if (this.DropOffLibraryDocuments.Count > 0)
                {
                    this.RoutingRules = rrm.RoutingRules;

                    RecordDocumentManager rdm = new RecordDocumentManager(
                        this.DropOffLibraryDocuments, rlm.EnabledLibraries, rrm.RoutingRules);
                    this.DropOffRecordDocuments = rdm.RecordDocuments;
                    rdm.ScanFields(rlm.DropOffLibrary, rlm.DropOffLibUrl);
                    rdm.MoveFileToLibrary();
                }
            }
        }
        #endregion

        #region DestructorDisposable
        /// <summary>
        /// Finalizes an instance of the <see cref="RecordCentreManager"/> class. 
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="RecordCentreManager"/> is reclaimed by garbage collection.
        /// </summary>
        ~RecordCentreManager()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttibutesOrProperties
        /// <summary>Gets the drop off library.</summary>
        public SPList DropOffLibrary { get; internal set; }

        /// <summary>Gets the enabled libraries.</summary>
        public List<SPList> EnabledLibraries { get; internal set; }

        /// <summary>Gets the routing rules.</summary>
        public SPListItemCollection RoutingRules { get; internal set; }

        /// <summary>Gets the drop off library records.</summary>
        public SPListItemCollection DropOffLibraryDocuments { get; internal set; }

        /// <summary>Gets the drop off record documents.</summary>
        public List<RecordDocument> DropOffRecordDocuments { get; internal set; }
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
            string s = "RecordsCount:\t\t" + this.DropOffLibraryDocuments.Count + "\n";
            s += "RecordsTitle:\t\t<";
            foreach (SPListItem spListItem in this.DropOffLibraryDocuments)
                s += spListItem.Title + ":" + spListItem.DisplayName + "|" + spListItem.File.Name + "?\n\t\t\t" +
                    spListItem.ContentType.Name + "|\n\t\t\t" +
                    spListItem.ContentType.Parent.Parent.Name + "||\n\t\t\t" +
                    spListItem.ContentType.Parent.Parent.Parent.Name + "|||\n\t\t\t";
            return s + ">\n";
        }
        #endregion

        #region PublicDisposable
        /// <summary>Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.</summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region PrivateDisposable
        /// <summary>Releases unmanaged and - optionally - managed resources.</summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// true to release both managed and unmanaged resources; false to release only unmanaged resources.
        /// <param name="isDisposing">The is disposing.</param>
        private void Dispose(bool isDisposing)
        {
            // Check if Dispose has been called
            if (!this.isDisposed)
            {
                // dispose managed and unmanaged resources
                if (isDisposing)
                {
                    // managed resources clean
                    this.DropOffLibrary = null;
                    this.EnabledLibraries = null;
                    this.DropOffLibraryDocuments = this.RoutingRules = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion
    }
}
