// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RecordDocument.cs" company="Montrium">
//   MIT License.
// </copyright>
// <summary>
//   
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using Microsoft.SharePoint;


    /// <summary>Record Manager.</summary>
    public class RecordDocument
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="RecordDocument"/> class.
        /// </summary>
        /// <param name="enabledLibraries">The enabled libraries.</param>
        /// <param name="fileListItem">The file sp list item.</param>
        /// <param name="routingRules">The routing rules.</param>
        public RecordDocument(List<SPList> enabledLibraries, SPListItem fileListItem, SPListItemCollection routingRules)
        {
            if (enabledLibraries == null) throw new ArgumentNullException("enabledLibraries");
            if (fileListItem == null) throw new ArgumentNullException("fileListItem");
            if (routingRules == null) throw new ArgumentNullException("routingRules");

            this.LogFile = Environment.GetEnvironmentVariable("USERPROFILE");
            this.LogFile = Path.Combine(this.LogFile, "Desktop");
            this.LogFile = Path.Combine(this.LogFile, fileListItem.File.ToString().Split('/')[1] + ".txt");
            this.ListItem = fileListItem;
            this.File = fileListItem.File;
            this.FieldCollection = fileListItem.Fields;
            this.ContentType = fileListItem.ContentType;
            this.ParentContentType = this.GetParentContentType(fileListItem.ContentType);
            this.HasRoutingRule = this.ValidateRoutingRule(routingRules);
            this.HasParentRoutingRule = this.ScreenForParentRoutingRule(routingRules);
            this.RoutingRule = this.GetRoutingRule(routingRules);
            this.ParentRoutingRule = this.GetParentRoutingRule(routingRules);
            this.HasLibrary = this.ValidateLibrary(enabledLibraries);
            this.HasParentLibrary = this.ValidateParentLibrary(enabledLibraries);
            this.CandidateLibrary = this.ScreenForLibrary(enabledLibraries);
            this.ParentCandidateLibrary = this.ScreenForParentLibrary(enabledLibraries);
            this.ScreenMetadata(fileListItem.Properties);
            this.TraceLog();
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="RecordDocument"/> class. Releases unmanaged resources and performs other cleanup operations before the <see cref="RecordDocument"/> is reclaimed by garbage collection.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class.
        ~RecordDocument()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets a value indicating whether this instance has library.</summary>
        public bool HasLibrary { get; internal set; }

        /// <summary>Gets a value indicating whether this instance has parent library.</summary>
        public bool HasParentLibrary { get; internal set; }

        /// <summary>Gets a value indicating whether this instance has routing rule.</summary>
        public bool HasRoutingRule { get; internal set; }

        /// <summary>Gets a value indicating whether this instance has parent routing rule.</summary>
        public bool HasParentRoutingRule { get; internal set; }

        /// <summary>Gets the routing rule.</summary>
        public SPListItem RoutingRule { get; internal set; }

        /// <summary>Gets the parent routing rule.</summary>
        public SPListItem ParentRoutingRule { get; internal set; }

        /// <summary>Gets the author.</summary>
        public string Author { get; internal set; }

        /// <summary>Gets the modified by.</summary>
        public string ModifiedBy { get; internal set; }

        /// <summary>Gets the title.</summary>
        public string Title { get; internal set; }

        /// <summary>Gets the content type id.</summary>
        public string ContentTypeId { get; internal set; }

        /// <summary>Gets the XML properties.</summary>
        public string XmlProperties { get; internal set; }

        /// <summary>Gets the log file.</summary>
        public string LogFile { get; internal set; }

        /// <summary>Gets the candidate library.</summary>
        public SPList CandidateLibrary { get; internal set; }

        /// <summary>Gets the parent candidate library.</summary>
        public SPList ParentCandidateLibrary { get; internal set; }

        /// <summary>Gets ListItem.</summary>
        public SPListItem ListItem { get; internal set; }

        /// <summary>Gets the file.</summary>
        public SPFile File { get; internal set; }

        /// <summary>Gets the field collection.</summary>
        public SPFieldCollection FieldCollection { get; internal set; }

        /// <summary>Gets ContentType.</summary>
        public SPContentType ContentType { get; internal set; }

        /// <summary>Gets Parent ContentType.</summary>
        public SPContentType ParentContentType { get; internal set; }
        #endregion

        #region PublicMethods
        /// <summary>
        /// Moves to library without folder.
        /// </summary>
        /// <param name="newUrl">The new URL.</param>
        public void MoveToLibraryWithoutFolder(string newUrl)
        {
            this.File.MoveTo(newUrl + this.File.Name, true);
        }

        /// <summary>
        /// Moves to library with folder.
        /// </summary>
        /// <param name="newFolder">
        /// The new Folder.
        /// </param>
        public void MoveToLibraryWithFolder(SPFolder newFolder)
        {
            this.File.MoveTo(newFolder.Url + "/" + this.File.Name, true);
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
            string s = this.Title.PadRight(20, '=');
            s += "\nLOG:\t\t\t\t" + this.LogFile + "\n";
            s += "SPListItem:\t\t\t" + this.ListItem + "\nContentTypeID:\t\t" + this.ContentTypeId + "\n";
            s += "Actual CTName:\t\t" + this.ContentType.Name + "\n";
            s += "HasRoutingRule:\t\t" + this.HasRoutingRule + "\n";
            s += "IsAdoptable?\t\t" + this.HasLibrary + "\n";
            s += "CandidateLib?\t\t" + this.CandidateLibrary + "\n";
            s += "Parent CTName:\t\t" + this.ParentContentType.Name + "\n";
            s += "HasParentRouting:\t" + this.HasParentRoutingRule + "\n";
            s += "IsParentAdoptable:\t" + this.HasParentLibrary + "\n";
            s += "ParentCandidateLib?\t" + this.ParentCandidateLibrary + "\n";
            return s;
        }

        /// <summary>
        /// Traces the log.
        /// </summary>
        private void TraceLog()
        {
            using (StreamWriter sr = new StreamWriter(this.LogFile, false, Encoding.UTF8))
                sr.WriteLine(this.ToString());
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

        #region PublicMethods
        /// <summary>
        /// Forces the content type to library.
        /// </summary>
        /// <param name="spContentType">The c type.</param>
        /// <param name="librarySpList">The library sp list.</param>
        public void ForceContentTypeToLibrary(SPContentType spContentType, SPList librarySpList)
        {
            if (spContentType == null) throw new ArgumentNullException(spContentType.Name);
            if (librarySpList == null) throw new ArgumentNullException(librarySpList.Title);

            // Add the content type to the list.
            /*if (!librarySpList.IsContentTypeAllowed(spContentType))
                MessageBox.Show("The " + spContentType.Name + " content type is not allowed on the " + librarySpList.Title + " list");*/
            /*else if (librarySpList.ContentTypes[spContentType.Name] != null)
                MessageBox.Show("The content type name " + spContentType.Name + " is already in use on the " + librarySpList.Title + " list");*/
            /*else
            {*/

                // if (librarySpLsit.ContentTypes[cType.Name] == null)
                librarySpList.ContentTypes.Add(spContentType);
                librarySpList.Update();
            /*}*/
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
                    this.Author = this.ModifiedBy = this.Title = this.ContentTypeId = this.XmlProperties = null;
                    this.CandidateLibrary = this.ParentCandidateLibrary = null;
                    this.ListItem = this.RoutingRule = this.ParentRoutingRule = null;
                    this.File = null;
                    this.FieldCollection = null;
                    this.ContentType = this.ParentContentType = null;
                }

                // unmanaged resources clean
                this.HasLibrary = this.HasParentLibrary = this.HasRoutingRule = this.HasParentRoutingRule = false;

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion

        #region PrivateMethods
        /// <summary>
        /// Gets the type of the parent content.
        /// </summary>
        /// <param name="spContentType">Type of the SP content.</param>
        /// <returns>The Parent SPContentType.</returns>
        private SPContentType GetParentContentType(SPContentType spContentType)
        {
            SPContentType aSpContentType = spContentType.Parent;
            while (aSpContentType.Name.Trim().Equals(spContentType.Name.Trim()))
                aSpContentType = aSpContentType.Parent;
            return aSpContentType;
        }

        /// <summary>Screens the metadata.</summary>
        /// <param name="hashtable">The hashtable.</param>
        private void ScreenMetadata(Hashtable hashtable)
        {
            foreach (DictionaryEntry diccoEntry in hashtable)
                if (diccoEntry.Key.Equals("_vti_RoutingExistingProperties"))
                    this.XmlProperties = diccoEntry.Value.ToString().Replace("<Value/>", "<Value></Value>").Replace("<Value />", "<Value></Value>");
                else if (diccoEntry.Key.Equals("vti_author")) this.Author = diccoEntry.Value.ToString();
                else if (diccoEntry.Key.Equals("vti_modifiedby")) this.ModifiedBy = diccoEntry.Value.ToString();
                else if (diccoEntry.Key.Equals("vti_title")) this.Title = diccoEntry.Value.ToString();
                else if (diccoEntry.Key.Equals("ContentTypeId")) this.ContentTypeId = diccoEntry.Value.ToString();
        }

        /// <summary>
        /// Validates the routing rule.
        /// </summary>
        /// <param name="spListItemCollection">The sp list item collection.</param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        private bool ValidateRoutingRule(SPListItemCollection spListItemCollection)
        {
            foreach (SPListItem spListItem in spListItemCollection)
                if (spListItem["RoutingContentType"].ToString().Trim().Contains(this.ContentType.Name))
                    return true;

            return false;
        }

        /// <summary>
        /// Gets the routing rule.
        /// </summary>
        /// <param name="spListItemCollection">
        /// The sp list item collection.
        /// </param>
        /// <returns>
        /// The Microsoft.SharePoint.SPListItem.
        /// </returns>
        private SPListItem GetRoutingRule(SPListItemCollection spListItemCollection)
        {
            foreach (SPListItem spListItem in spListItemCollection)
                if (spListItem["RoutingContentType"].ToString().Trim().Contains(this.ContentType.Name))
                    return spListItem;

            return null;
        }

        /// <summary>
        /// Gets the parent routing rule.
        /// </summary>
        /// <param name="spListItemCollection">
        /// The sp list item collection.
        /// </param>
        /// <returns>
        /// The Microsoft.SharePoint.SPListItem.
        /// </returns>
        private SPListItem GetParentRoutingRule(SPListItemCollection spListItemCollection)
        {
            foreach (SPListItem spListItem in spListItemCollection)
                if (spListItem["RoutingContentType"].ToString().Trim().Contains(this.ParentContentType.Name))
                    return spListItem;

            return null;
        }

        /// <summary>
        /// Screens for parent routing rule.
        /// </summary>
        /// <param name="spListItemCollection">The sp list item collection.</param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        private bool ScreenForParentRoutingRule(SPListItemCollection spListItemCollection)
        {
            foreach (SPListItem spListItem in spListItemCollection)
                if (spListItem["RoutingContentType"].ToString().Trim().Contains(this.ParentContentType.Name))
                    return true;

            return false;
        }

        /// <summary>
        /// Screens for library.
        /// </summary>
        /// <param name="spListCollection">
        /// The sp list collection.
        /// </param>
        /// <returns>
        /// The Microsoft.SharePoint.SPList.
        /// </returns>
        private SPList ScreenForLibrary(List<SPList> spListCollection)
        {
            foreach (SPList librarySpList in spListCollection)
                if (librarySpList.ContentTypes[this.ContentType.Name] != null
                    && librarySpList.ContentTypesEnabled)
                    return librarySpList;

            return null;
        }

        /// <summary>
        /// Screens for parent library.
        /// </summary>
        /// <param name="spListCollection">
        /// The sp list collection.
        /// </param>
        /// <returns>
        /// The Microsoft.SharePoint.SPList.
        /// </returns>
        private SPList ScreenForParentLibrary(List<SPList> spListCollection)
        {
            foreach (SPList librarySpList in spListCollection)
                if (librarySpList.ContentTypes[this.ParentContentType.Name] != null
                    && librarySpList.ContentTypesEnabled)
                    return librarySpList;

            return null;
        }

        /// <summary>
        /// Validates the library.
        /// </summary>
        /// <param name="spListCollection">The sp List Collection.</param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        private bool ValidateLibrary(IEnumerable<SPList> spListCollection)
        {
            bool boo = false;

            foreach (SPList librarySpList in spListCollection)
                if (librarySpList.ContentTypes[this.ContentType.Name] != null)
                {
                    boo = true;
                    try
                    {
                        librarySpList.ContentTypesEnabled = true;
                        boo = true;
                    }
                    catch
                    {
                        boo = false;
                    }
                }

            return boo;
        }

        /// <summary>
        /// Validates the parent library.
        /// </summary>
        /// <param name="spListCollection">The sp list collection.</param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        private bool ValidateParentLibrary(IEnumerable<SPList> spListCollection)
        {
            bool boo = false;

            foreach (SPList librarySpList in spListCollection)
                if (librarySpList.ContentTypes[this.ParentContentType.Name] != null)
                {
                    boo = true;
                    try
                    {
                        librarySpList.ContentTypesEnabled = true;
                        boo = true;
                    }
                    catch
                    {
                        boo = false;
                    }
                }

            return boo;
        }
        #endregion
    }
}
