// -----------------------------------------------------------------------
// <copyright file="ContentOrganizerRuler.cs" company="Montrium">
// MIT License.
// </copyright>
// -----------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.Windows.Forms;

    using Microsoft.SharePoint.Client;

    /// <summary>ContentOrganizer Ruler.</summary>
    public class ContentOrganizerRuler : IDisposable
    {
        #region fields
        /// <summary>Is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructor
        /// <summary>Initializes a new instance of the <see cref="ContentOrganizerRuler"/> class.</summary>
        /// <param name="rulesListItemCollection">The rules sp list collection.</param>
        public ContentOrganizerRuler(ListItemCollection rulesListItemCollection)
        {
            // replace with your upload content type ID.
            const string defaultContentTypeId = "0x01010B";
            ListItem rule = null;
            string contentType = String.Empty;

            foreach (ListItem listItem in rulesListItemCollection)
            {
                contentType = String.Empty;
                string contentTypeId = String.Empty;

                if (listItem.FieldValues.ContainsKey("RoutingContentTypeInternal"))
                {
                    var value = listItem.FieldValues["RoutingContentTypeInternal"] ?? String.Empty;
                    string[] values = value.ToString().Split("|".ToCharArray(), StringSplitOptions.None);

                    if (values.Length == 2)
                    {
                        contentTypeId = values[0];
                        contentType = values[1];
                    }
                }

                if (defaultContentTypeId == contentTypeId)
                {
                    rule = listItem;
                    break;
                }
            }

            MessageBox.Show(rule != null ? "Send to Drop Off Library" : "Send to Content Type Library" + contentType);
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="ContentOrganizerRuler"/> class.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~ContentOrganizerRuler()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
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
            return base.ToString();
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
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion
    }
}
