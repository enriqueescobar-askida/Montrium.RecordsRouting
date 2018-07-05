// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XmlLookupNode.cs" company="Montrium">
//   MIT License
// </copyright>
// <summary>
//   Defines the XmlLookupNode type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using Microsoft.SharePoint;

    /// <summary>The XML Lookup Node.</summary>
    public class XmlLookupNode : IDisposable
    {
        #region fields
        /// <summary>Boolean isDisposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="XmlLookupNode"/> class.</summary>
        /// <param name="fieldName">The field name.</param>
        /// <param name="fieldValue">The field value.</param>
        /// <param name="fieldType">Type of the field.</param>
        public XmlLookupNode(string fieldName, string fieldValue, string fieldType)
        {
            if (String.IsNullOrEmpty(fieldName) || String.IsNullOrEmpty(fieldType))
                throw new ArgumentNullException("fieldName", "fieldName or fieldtype NULL");

            this.CamelCaseName = fieldName.Split(';')[0].Replace("_x0020_", "").Replace("xd_", "").Replace("_", "");
            this.Type = fieldType;
            this.Value = fieldValue;
        }
        #endregion

        #region DestructorDisposable
        /// <summary>
        /// Finalizes an instance of the <see cref="XmlLookupNode"/> class.
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// <see cref="XmlLookupNode"/> is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class.
        /// </summary>
        ~XmlLookupNode()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets the camel case name.</summary>
        public string CamelCaseName { get; internal set; }

        /// <summary>Gets the type of the field.</summary>
        public string Value { get; internal set; }

        /// <summary>Gets the type of the field.</summary>
        public string Type { get; internal set; }
        #endregion

        #region PublicMethods
        /// <summary>
        /// Gets the name of the sentence.
        /// </summary>
        /// <returns>
        /// The System.String.
        /// </returns>
        public string GetSentenceName()
        {
            string anyString = this.CamelCaseName;
            if (String.IsNullOrEmpty(anyString)) return String.Empty;

            anyString = Char.ToUpperInvariant(anyString[0]) + anyString.Substring(1);
            return System.Text.RegularExpressions.Regex.Replace(anyString, "([^^])([A-Z])", "$1 $2");
        }

        /// <summary>
        /// Gets the type of the sp field.
        /// </summary>
        /// <returns>
        /// The Microsoft.SharePoint.SPFieldType.
        /// </returns>
        public SPFieldType GetSpFieldType()
        {
            if (this.Type == "AllDayEvent")
                return SPFieldType.AllDayEvent;
            else if (this.Type == "Attachments")
                return SPFieldType.Attachments;
            else if (this.Type == "Boolean")
                return SPFieldType.Boolean;
            else if (this.Type == "Calculated")
                return SPFieldType.Calculated;
            else if (this.Type == "Choice")
                return SPFieldType.Choice;
            else if (this.Type == "Computed")
                return SPFieldType.Computed;
            else if (this.Type == "ContentTypeId")
                return SPFieldType.ContentTypeId;
            else if (this.Type == "Counter")
                return SPFieldType.Counter;
            else if (this.Type == "CrossProjectLink")
                return SPFieldType.CrossProjectLink;
            else if (this.Type == "Currency")
                return SPFieldType.Currency;
            else if (this.Type == "DateTime")
                return SPFieldType.DateTime;
            else if (this.Type == "Error")
                return SPFieldType.Error;
            else if (this.Type == "File")
                return SPFieldType.File;
            else if (this.Type == "GridChoice")
                return SPFieldType.GridChoice;
            else if (this.Type == "Guid")
                return SPFieldType.Guid;
            else if (this.Type == "Integer")
                return SPFieldType.Integer;
            else if (this.Type == "Invalid")
                return SPFieldType.Invalid;
            else if (this.Type == "Lookup")
                return SPFieldType.Lookup;
            else if (this.Type == "MaxItems")
                return SPFieldType.MaxItems;
            else if (this.Type == "ModStat")
                return SPFieldType.ModStat;
            else if (this.Type == "MultiChoice")
                return SPFieldType.MultiChoice;
            else if (this.Type == "Note")
                return SPFieldType.Note;
            else if (this.Type == "Number")
                return SPFieldType.Number;
            else if (this.Type == "PageSeparator")
                return SPFieldType.PageSeparator;
            else if (this.Type == "Recurrence")
                return SPFieldType.Recurrence;
            else if (this.Type == "Text")
                return SPFieldType.Text;
            else if (this.Type == "ThreadIndex")
                return SPFieldType.ThreadIndex;
            else if (this.Type == "Threading")
                return SPFieldType.Threading;
            else if (this.Type == "URL")
                return SPFieldType.URL;
            else if (this.Type == "User")
                return SPFieldType.User;
            else if (this.Type == "WorkflowEventType")
                return SPFieldType.WorkflowEventType;
            else if (this.Type == "WorkflowStatus")
                return SPFieldType.WorkflowStatus;
            else return SPFieldType.Lookup;
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
            return "FieldName:\t" + this.CamelCaseName + "\nFieldValue:\t" + this.Value +
                "\nFieldType:\t" + this.Type + "\n";
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
                    this.CamelCaseName = this.Value = this.Type = String.Empty;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion
    }
}