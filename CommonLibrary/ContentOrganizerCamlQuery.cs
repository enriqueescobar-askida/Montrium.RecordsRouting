// -----------------------------------------------------------------------
// <copyright file="ContentOrganizerCamlQuery.cs" company="Montrium">
// MIT License
// </copyright>
// -----------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.Xml.Linq;

    using Microsoft.SharePoint.Client;

    /// <summary>ContentOrganizer Caml Query.</summary>
    public class ContentOrganizerCamlQuery : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;

        /// <summary>view Fields.</summary>
        private string[] viewFields = new string[]
               {
                   "RoutingConditions",
                   "RoutingContentTypeInternal",
                   "RoutingPriority",
                   "RoutingRuleName",
                   "RoutingTargetFolder",
                   "RoutingTargetLibrary",
                   "RoutingTargetPath"
               };
        #endregion

        #region Constructor
        /// <summary>Initializes a new instance of the <see cref="ContentOrganizerCamlQuery"/> class.</summary>
        public ContentOrganizerCamlQuery()
        {
            // view...
            CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery(100, this.ViewFields);
            XElement xElementView = XElement.Parse(camlQuery.ViewXml);

            // query...
            XElement xElementRoutingEnabled =
                new XElement(
                            "Eq",
                            new XElement("FieldRef", new XAttribute("Name", "RoutingEnabled")),
                            new XElement("Value", new XAttribute("Type", "YesNo"), "1"));

            XElement xElementQuery = new XElement("Query", new XElement("Where", xElementRoutingEnabled));

            // Add query element to view element
            xElementView.FirstNode.AddBeforeSelf(xElementQuery);
            camlQuery.ViewXml = xElementView.ToString();
            this.CamlQueryIt = camlQuery;
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="ContentOrganizerCamlQuery"/> class.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~ContentOrganizerCamlQuery()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets the view fields.</summary>
        public string[] ViewFields
        {
            get { return this.viewFields; }
            internal set { this.viewFields = value; }
        }

        /// <summary>Gets the caml query it.</summary>
        public CamlQuery CamlQueryIt { get; internal set; }
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
                    this.CamlQueryIt = null;
                    this.ViewFields = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion
    }
}
