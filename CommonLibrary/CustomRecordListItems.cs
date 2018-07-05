// -----------------------------------------------------------------------
// <copyright file="CustomRecordList.cs" company="Montrium">
// MIT License
// </copyright>
// -----------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.Collections;

    using Microsoft.Office.RecordsManagement.RecordsRepository;
    using Microsoft.SharePoint;

    /// <summary>Custom RecordListItems.</summary>
    public class CustomRecordListItems : IDisposable
    {
        #region Constructors

        /// <summary>Initializes a new instance of the <see cref="CustomRecordListItems"/> class.</summary>
        /// <param name="url">The url.</param>
        /// <param name="sourceFolder">The source folder.</param>
        /// <param name="destinationFolder">The destination folder.</param>
        public CustomRecordListItems(string url, string sourceFolder, string destinationFolder)
        {
            if (String.IsNullOrEmpty(url))
                throw new ArgumentNullException(url, "Inavild url");

            if (String.IsNullOrEmpty(sourceFolder))
                throw new ArgumentNullException(sourceFolder, "Invalid folder");

            if (String.IsNullOrEmpty(destinationFolder))
                throw new ArgumentNullException(destinationFolder, "Invalid folder");

            this.DocumentsSourceFolder = sourceFolder;
            this.DocumentsFinalFolder = destinationFolder;

            using (SPSite spSite = new SPSite(url))
            using (SPWeb spWeb = spSite.OpenWeb())
            {
                // EcmDoc
                EcmDocumentRoutingWeb ecmDocumentRoutingWeb = new EcmDocumentRoutingWeb(spWeb);

                // current user specs
                this.LoginName = System.Threading.Thread.CurrentPrincipal.Identity.Name;
                this.DocumentsUser = SPContext.Current.Web.CurrentUser;

                // screening the source folder
                SPList spList = spWeb.Lists[sourceFolder];
                this.DocumentsFields = spList.Fields;
                int fieldsLimit = spList.Fields.Count;

                // viewing all items in source folder
                SPView spView = spList.Views["All Items"];

                // current view fields on all items in source folder
                this.DocumentsViewFields = spView.ViewFields;
                int viewFieldsLimit = spView.ViewFields.Count;

                // screening the item collection for each item
                this.DocumentsItems = spList.Items;

                foreach (SPListItem spListItem in this.DocumentsItems)
                {
                    SPContentType spContentType = spListItem.ContentType;
                    SPCopyFieldMask spCopyFieldMask = spListItem.CopyFieldMask;
                    SPCopyDestinationCollection spCopyDestinationCollection = spListItem.CopyDestinations;
                    string displayName = spListItem.DisplayName;
                    SPBasePermissions spBasePermissions = spListItem.EffectiveBasePermissions;
                    SPFieldCollection spFieldCollection = spListItem.Fields;
                    SPFile spFile = spListItem.File;
                    SPFolder spFolder = spListItem.Folder;
                    SPFileLevel spFileLevel = spListItem.Level;
                    string name = spListItem.Name;
                    Hashtable properties = spListItem.Properties;
                    string title = spListItem.Title;
                    string strUrl = spListItem.Url;
                    SPListItemVersionCollection spListItemVersionCollection = spListItem.Versions;
                    string xml = spListItem.Xml;
                }
            }
        }
        #endregion

        #region DestructorDispose
        /// <summary>
        /// Finalizes an instance of the <see cref="CustomRecordListItems"/> class.
        /// </summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~CustomRecordListItems()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region Attibutes
        /// <summary>Is Disposed.</summary>
        private bool IsDisposed = false;

        /// <summary>Gets the documents URL.</summary>
        public string DocumentsUrl { get; internal set; }

        /// <summary>Gets the documents source folder.</summary>
        public string DocumentsSourceFolder { get; internal set; }

        /// <summary>Gets the documents final folder.</summary>
        public string DocumentsFinalFolder { get; internal set; }

        /// <summary>Gets LoginName.</summary>
        public string LoginName { get; internal set; }

        /// <summary>Gets the documents items.</summary>
        public SPListItemCollection DocumentsItems { get; internal set; }

        /// <summary>Gets DocumentsFields.</summary>
        public SPFieldCollection DocumentsFields { get; internal set; }

        /// <summary>Gets the documents view fields.</summary>
        public SPViewFieldCollection DocumentsViewFields { get; internal set; }

        /// <summary>Gets the documents user.</summary>
        public SPUser DocumentsUser { get; internal set; }
        #endregion

        #region DisposablePrivateMethods
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region DisposablePublicMethods
        /// <summary>Dispose Finally.
        /// </summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// true to release both managed and unmanaged resources; false to release only unmanaged resources.
        /// <param name="isDisposing">
        /// The is disposing.
        /// </param>
        private void Dispose(bool isDisposing)
        {
            // Check if Dispose has been called
            if (!this.IsDisposed)
            {
                // dispose managed and unmanaged resources
                if (isDisposing)
                {
                    // managed resources clean
                    this.DocumentsSourceFolder = this.DocumentsFinalFolder = null;
                    this.DocumentsUrl = null;
                    this.LoginName = null;
                    this.DocumentsUser = null;
                    this.DocumentsViewFields = null;
                    this.DocumentsItems = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.IsDisposed = true;
            }
        }
        #endregion
    }
}
