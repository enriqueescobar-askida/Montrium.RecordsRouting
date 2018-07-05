// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XmlLookupReader.cs" company="Montrium">
//   MIT License
// </copyright>
// <summary>
//   Defines the XmlLookupReader type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.SharePoint;

    /// <summary>XmlLookup Reader.</summary>
    public class XmlLookupReader : IDisposable
    {
        #region fields
        /// <summary>If this is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="XmlLookupReader"/> class.</summary>
        /// <param name="xmlProperties">The xml properties.</param>
        public XmlLookupReader(string xmlProperties)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xmlProperties);
            xmlDocument.LoadXml(xmlDocument.FirstChild.FirstChild.OuterXml);
            this.LookupNodeList = this.FindXmlLookups(xmlDocument.DocumentElement);
        }
        #endregion

        #region DestructorDisposable
        /// <summary>
        /// Finalizes an instance of the <see cref="XmlLookupReader"/> class.
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="XmlLookupReader"/> is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class.
        /// </summary>
        ~XmlLookupReader()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets the XML lookup node list.</summary>
        public List<XmlLookupNode> LookupNodeList { get; internal set; }

        /// <summary>Gets the lookup node matched.</summary>
        public XmlLookupNode LookupNodeMatched { get; internal set; }
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
            return "LookupNodeList.Count:\t" + this.LookupNodeList.Count + "\nXmlLookupNode:\t" + this.LookupNodeMatched;
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
        /// <summary>Determines whether [contains] [the specified sp field].</summary>
        /// <param name="spField">The sp field.</param>
        /// <returns>
        ///   <c>true</c> if [contains] [the specified sp field]; otherwise, <c>false</c>.
        /// </returns>
        public bool Contains(SPField spField)
        {
            string fieldName = spField.Title.Replace(" ID", " Id").Replace("(Converted Document)", "").Replace(" ", String.Empty);
            List<XmlLookupNode> xmlLookupNodes = this.LookupNodeList;

            foreach (XmlLookupNode xmlLookupNode in xmlLookupNodes)
                if (xmlLookupNode.CamelCaseName.Equals(fieldName)) // && lookupNode.FieldType.Equals(spField.TypeAsString))
                {
                    this.LookupNodeMatched = xmlLookupNode;
                    return true;
                }

            return false;
        }

        /// <summary>
        /// Values the specified sp field.
        /// </summary>
        /// <param name="spField">
        /// The sp field.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        public string Value(SPField spField)
        {
            string fieldName =
                spField.Title.Replace(" ID", " Id").Replace("(Converted Document)", "").Replace(" ", String.Empty);
            List<XmlLookupNode> xmlLookupNodes = this.LookupNodeList;

            foreach (XmlLookupNode xmlLookupNode in xmlLookupNodes)
                if (xmlLookupNode.CamelCaseName.Equals(fieldName)) // && lookupNode.FieldType.Equals(spField.TypeAsString))
                    return xmlLookupNode.Value;

            return String.Empty;
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
                    this.LookupNodeList = null;
                    this.LookupNodeMatched = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion

        #region PrivateMethods
        /// <summary>
        /// Finds the lookups.
        /// </summary>
        /// <param name="xmlDocumentElement">
        /// The xml Document Element.
        /// </param>
        /// <returns>
        /// The System.Collections.Generic.List`1[T -&gt; Mtm.RecordsRouting.XmlLookupNode].
        /// </returns>
        private List<XmlLookupNode> FindXmlLookups(XmlElement xmlDocumentElement)
        {
            List<XmlLookupNode> xmlLookupNodesList = new List<XmlLookupNode>();

            /* go to property list skipping:
             * null values,
             * more than 3 children,
             * without {MetaInfo,vit,displayurn}* name,
             * without Computed type
             */
            foreach (XmlNode xmlNodeProperty in xmlDocumentElement.ChildNodes)
                if (xmlNodeProperty.ChildNodes.Count == 3 &&
                    !String.IsNullOrEmpty(xmlNodeProperty.FirstChild.NextSibling.InnerText) &&
                    this.IsInternalNameUse(xmlNodeProperty.FirstChild.InnerText) &&
                    this.IsTitleValid(xmlNodeProperty.FirstChild.InnerText) &&
                    this.IsInternalNameValid(xmlNodeProperty.FirstChild.InnerText) &&
                    this.IsIdValid(xmlNodeProperty.FirstChild.InnerText) &&
                    this.IsTypeValid(xmlNodeProperty.LastChild.InnerText))
                    xmlLookupNodesList.Add(
                        new XmlLookupNode(xmlNodeProperty.FirstChild.InnerText,
                            xmlNodeProperty.FirstChild.NextSibling.InnerText,
                            xmlNodeProperty.LastChild.InnerText));

            return xmlLookupNodesList;
        }

        /// <summary>
        /// Determines whether [is internal name use] [the specified name].
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>
        ///   <c>true</c> if [is internal name use] [the specified name]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsInternalNameUse(string name)
        {
            name = name.Trim();
            return !name.StartsWith("vti") && !name.ToLowerInvariant().Contains("display");
        }

        /// <summary>
        /// Determines whether [is title valid] [the specified title].
        /// </summary>
        /// <param name="title">The title.</param>
        /// <returns>
        ///   <c>true</c> if [is title valid] [the specified title]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsTitleValid(string title)
        {
            title = title.Trim();
            return !title.Equals("Signatures Status") && !title.Equals("SignaturesStatus") && !title.Equals("Name");
        }

        /// <summary>
        /// Determines whether [is internal name valid] [the specified sp field internal name].
        /// </summary>
        /// <param name="spFieldInternalName">Name of the sp field internal.</param>
        /// <returns>
        ///   <c>true</c> if [is internal name valid] [the specified sp field internal name]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsInternalNameValid(string spFieldInternalName)
        {
            spFieldInternalName = spFieldInternalName.Trim();
            return !spFieldInternalName.Contains("TaxCatchAll") && !spFieldInternalName.Contains("TaxCatchAllLabel");
        }

        /// <summary>
        /// Determines whether [is type valid] [the specified type].
        /// </summary>
        /// <param name="spFieldType">The type.</param>
        /// <returns>
        ///   <c>true</c> if [is type valid] [the specified type]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsTypeValid(string spFieldType)
        {
            spFieldType = spFieldType.Trim();
            return !spFieldType.Equals("Attachments") && !spFieldType.Equals("File") && !spFieldType.Equals("Computed");
        }

        /// <summary>
        /// Determines whether [is id valid] [the specified type].
        /// </summary>
        /// <param name="guid">The GUID.</param>
        /// <returns>
        ///   <c>true</c> if [is id valid] [the specified type]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsIdValid(string guid)
        {
            guid = guid.Trim();
            return !guid.Contains("DocIcon") // "Type"
                    && !guid.Contains("ContentTypeId") // "ContentTypeId"
                    && !guid.Contains("ContentType") // "ContentType"
                    && !guid.Contains("TemplateUrl") // "Template Link"
                    && !guid.Contains("xd_ProgID") // "Html File Link"
                    && !guid.Contains("xd_Signature") // "Is Signed"
                    && !guid.Contains("MetaInfo"); // "Property Bag"
        }
        #endregion
    }
}