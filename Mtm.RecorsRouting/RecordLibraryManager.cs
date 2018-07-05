// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RecordLibraryManager.cs" company="Montrium">
//   MIT License.
// </copyright>
// <summary>
//   
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using Microsoft.Office.RecordsManagement.RecordsRepository;
    using Microsoft.SharePoint;


    /// <summary>ContentType Library.</summary>
    public class RecordLibraryManager : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="RecordLibraryManager"/> class.</summary>
        /// <param name="url">The URL.</param>
        public RecordLibraryManager(string url)
        {
            using (SPSite spSite = new SPSite(url))
            using (SPWeb spWeb = spSite.OpenWeb())
            {
                spWeb.AllowUnsafeUpdates = true;
                EcmDocumentRoutingWeb edrw = new EcmDocumentRoutingWeb(spWeb);
                this.DropOffLibUrl = edrw.DropOffZoneUrl;
                this.DropOffLibTitle = edrw.DropOffZone.Title;
                this.DropOffLibrary = spWeb.Lists.TryGetList(edrw.DropOffZone.Title);
                this.EnabledLibraries = this.ScreenLibraries(spWeb);
            }
        }
        #endregion

        #region DestructorDispose
        /// <summary>Finalizes an instance of the <see cref="RecordLibraryManager"/> class.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~RecordLibraryManager()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets the drop off lib URL.</summary>
        public string DropOffLibUrl { get; internal set; }

        /// <summary>Gets the drop off lib title.</summary>
        public string DropOffLibTitle { get; internal set; }

        /// <summary>Gets the drop off library.</summary>
        public SPList DropOffLibrary { get; internal set; }

        /// <summary>Gets the enabled libraries.</summary>
        public List<SPList> EnabledLibraries { get; internal set; } 
        #endregion

        #region PublicMethods
        /// <summary>
        /// Gets the content type boolean.
        /// </summary>
        /// <param name="recordDocument">
        /// The record document.
        /// </param>
        /// <param name="checkParent">
        /// The check parent.
        /// </param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        public bool GetContentTypeBoolean(RecordDocument recordDocument, bool checkParent)
        {
            if (!checkParent)
                return this.GetContentTypeBoolean(recordDocument);
            else
            {
                bool boo = false;

                foreach (SPList enabledLibrary in this.EnabledLibraries)
                    if (enabledLibrary.ContentTypesEnabled &&
                        enabledLibrary.ContentTypes[recordDocument.ParentContentType.Name] != null)
                        boo = true;

                return this.GetContentTypeBoolean(recordDocument) || boo;
            }
        }

        /// <summary>
        /// Gets the content type indices.
        /// </summary>
        /// <param name="recordDocument">
        /// The record document.
        /// </param>
        /// <param name="checkParent">
        /// if set to <c>true</c> [check parent].
        /// </param>
        /// <returns>
        /// The System.Collections.Generic.List`1[T -&gt; System.Int32].
        /// </returns>
        public List<int> GetContentTypeIndices(RecordDocument recordDocument, bool checkParent)
        {
            if (!checkParent) return this.GetContentTypeIndices(recordDocument);
            else
            {
                List<int> intList = new List<int>();

                for (int i = 0; i < this.EnabledLibraries.Count; i++)
                {
                    SPList enabledLibrary = this.EnabledLibraries[i];
                    if (enabledLibrary.ContentTypesEnabled &&
                        enabledLibrary.ContentTypes[recordDocument.ParentContentType.Name] != null)
                        intList.Add(i);
                }

                return intList;
            }
        }

        /// <summary>
        /// Prints to log file.
        /// </summary>
        /// <param name="logPath">
        /// The log Path.
        /// </param>
        public void PrintToLogFile(string logPath)
        {
            using (StreamWriter sw = new StreamWriter(logPath.Replace(".txt", "Libraries.txt"), false, Encoding.UTF8))
            {
                sw.WriteLine(this.ToString());
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
            string s = "DoL_url:\t" + this.DropOffLibUrl + "\nDoL_tit:\t" + this.DropOffLibTitle + "\n";
            s += "DoL_libc:\t" + this.DropOffLibrary.ItemCount + "\n";
            foreach (SPList enabledLibrary in this.EnabledLibraries)
            {
                s += enabledLibrary.BaseType.ToString().PadRight(30, '=') + enabledLibrary.Title + "\n";
                s += enabledLibrary.Author.Name + "\n";
                s += enabledLibrary.ContentTypesEnabled + "\n";
                s += enabledLibrary.EnableFolderCreation + "\n";
                s += enabledLibrary.Created.ToLocalTime() + "\n";
                s += enabledLibrary.Description + "\n";
                s += enabledLibrary.Fields.Count + "\n";
                s += enabledLibrary.Folders.Count + "\n";
                s += enabledLibrary.ItemCount + "\n";
                s += enabledLibrary.Items.Count + "\n";
                s += enabledLibrary.ParentWeb.Url + "\n";
                s += enabledLibrary.SendToLocationName + "\n";
                s += enabledLibrary.SendToLocationUrl + "\n";
                s += enabledLibrary.ContentTypes[0].Name + "\n";
                s += enabledLibrary.GetType().Name + "\n";
            }

            return s + "\n";
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
                    this.DropOffLibUrl = this.DropOffLibTitle = String.Empty;
                    this.DropOffLibrary = null;
                    this.EnabledLibraries = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion

        #region PrivateMethods
        /// <summary>
        /// Gets the content type boolean.
        /// </summary>
        /// <param name="recordDocument">The record document.</param>
        /// <returns>
        ///   <c>true</c> if [has content type] [the specified record document]; otherwise, <c>false</c>.
        /// </returns>
        private bool GetContentTypeBoolean(RecordDocument recordDocument)
        {
            foreach (SPList enabledLibrary in this.EnabledLibraries)
                if (enabledLibrary.ContentTypesEnabled &&
                    enabledLibrary.ContentTypes[recordDocument.ContentType.Name] != null)
                    return true;

            return false;
        }

        /// <summary>
        /// Gets the content type indices.
        /// </summary>
        /// <param name="recordDocument">
        /// The record document.
        /// </param>
        /// <returns>
        /// The System.Collections.Generic.List`1[T -&gt; System.Int32].
        /// </returns>
        private List<int> GetContentTypeIndices(RecordDocument recordDocument)
        {
            List<int> intList = new List<int>();
            for (int i = 0; i < this.EnabledLibraries.Count; i++)
            {
                SPList enabledLibrary = this.EnabledLibraries[i];
                if (enabledLibrary.ContentTypesEnabled &&
                    enabledLibrary.ContentTypes[recordDocument.ContentType.Name] != null)
                    intList.Add(i);
            }

            return intList;
        }

        /// <summary>
        /// Screens the libraries.
        /// </summary>
        /// <param name="spWeb">
        /// The sp Web.
        /// </param>
        /// <returns>
        /// The System.Collections.Generic.List`1[T -&gt; Microsoft.SharePoint.SPList].
        /// </returns>
        private List<SPList> ScreenLibraries(SPWeb spWeb)
        {
            List<SPList> spLibraryList = new List<SPList>();

            foreach (SPList librarySpList in spWeb.Lists)
            {
                if ((librarySpList is SPPictureLibrary || librarySpList is SPDocumentLibrary)
                    && librarySpList.Title != this.DropOffLibTitle)
                {
                    try
                    {
                        librarySpList.ContentTypesEnabled = true;
                    }
                    catch
                    {
                    }

                    if (librarySpList.ContentTypesEnabled)
                        spLibraryList.Add(librarySpList);
                }
            }

            return spLibraryList;
        }

        /// <summary>
        /// The add library.
        /// </summary>
        /// <param name="spWeb">
        /// The sp Web.
        /// </param>
        /// <param name="libraryName">
        /// The library Name.
        /// </param>
        /// <param name="libraryDesc">
        /// The library Desc.
        /// </param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        private bool AddLibrary(SPWeb spWeb, string libraryName, string libraryDesc)
        {
            if (spWeb == null) throw new ArgumentNullException(spWeb.ToString());
            if (String.IsNullOrEmpty(libraryName)) throw new ArgumentNullException(libraryName);
            if (String.IsNullOrEmpty(libraryDesc)) throw new ArgumentNullException(libraryDesc);

            int count = spWeb.Lists.Count;
            spWeb.Lists.Add(libraryName, libraryDesc, SPListTemplateType.DocumentLibrary);
            spWeb.Update();

            return spWeb.Lists.Count == count + 1;
        }
        #endregion
    }
}