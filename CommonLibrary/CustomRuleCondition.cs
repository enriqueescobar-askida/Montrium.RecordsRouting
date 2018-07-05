// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CustomRuleCondition.cs" company="Montrium">
//   MIT Licence
// </copyright>
// <summary>
//   
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;

    using Microsoft.SharePoint;

    /// <summary>Custom RecordConditions.</summary>
    public class CustomRuleCondition : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="CustomRuleCondition"/> class.</summary>
        /// <param name="spField">The sp field.</param>
        /// <param name="strOperator">The str operator.</param>
        /// <param name="fieldValue">The field value.</param>
        public CustomRuleCondition(SPField spField, string strOperator, string fieldValue)
        {
            if (spField == null) throw new ArgumentNullException(spField.ToString());
            if (String.IsNullOrEmpty(strOperator)) throw new ArgumentNullException(strOperator);
            if (String.IsNullOrEmpty(fieldValue)) throw new ArgumentNullException(fieldValue);

            this.ConditionFieldId = spField.Id.ToString();
            this.ConditionFieldInternalName = spField.InternalName;
            this.ConditionFieldTitle = spField.Title;
            this.ConditionOperator = strOperator;
            this.ConditionFieldValue = fieldValue;

            this.XmlBody();
        }

        /// <summary>Initializes a new instance of the <see cref="CustomRuleCondition"/> class.</summary>
        /// <param name="fieldId">The field id.</param>
        /// <param name="fieldInternalName">The field internal name.</param>
        /// <param name="fieldTitle">The field title.</param>
        /// <param name="strOperator">The str operator.</param>
        /// <param name="fieldValue">The field value.</param>
        public CustomRuleCondition(string fieldId, string fieldInternalName, string fieldTitle,
            string strOperator, string fieldValue)
        {
            if (String.IsNullOrEmpty(fieldId)) throw new ArgumentNullException(fieldId);
            if (String.IsNullOrEmpty(fieldInternalName)) throw new ArgumentNullException(fieldInternalName);
            if (String.IsNullOrEmpty(fieldTitle)) throw new ArgumentNullException(fieldTitle);
            if (String.IsNullOrEmpty(strOperator)) throw new ArgumentNullException(strOperator);
            if (String.IsNullOrEmpty(fieldValue)) throw new ArgumentNullException(fieldValue);

            this.ConditionFieldId = fieldId;
            this.ConditionFieldInternalName = fieldInternalName;
            this.ConditionFieldTitle = fieldTitle;
            this.ConditionOperator = strOperator;
            this.ConditionFieldValue = fieldValue;

            this.XmlBody();
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="CustomRuleCondition"/> class.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~CustomRuleCondition()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region Attributes
        /// <summary>Gets or sets ConditionFieldId.</summary>
        public string ConditionFieldId { get; internal set; }

        /// <summary>Gets or sets ConditionFieldInternalName.</summary>
        public string ConditionFieldInternalName { get; internal set; }

        /// <summary>Gets or sets ConditionFieldTitle.</summary>
        public string ConditionFieldTitle { get; internal set; }

        /// <summary>Gets or sets ConditionOperator.</summary>
        public string ConditionOperator { get; internal set; }

        /// <summary>Gets or sets ConditionFieldValue.</summary>
        public string ConditionFieldValue { get; internal set; }

        /// <summary>Gets the XML conditions.</summary>
        public string XmlConditions { get; internal set; }
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
            string s = String.Empty;
            return s + "\n";
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
        /// <summary>Dispose Finally.</summary>
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
                    this.ConditionFieldId = this.ConditionFieldInternalName = null;
                    this.ConditionFieldTitle = this.ConditionOperator = null;
                    this.ConditionFieldValue = this.XmlConditions = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion

        #region PrivateMethods
        /// <summary>XMLs the body.</summary>
        private void XmlBody()
        {
            string conditionXmlBody =
                String.Format(
                    @"<Condition Column=""{0}|{1}|{2}"" Operator=""{3}"" Value=""{4}"" />",
                    this.ConditionFieldId,
                    this.ConditionFieldInternalName,
                    this.ConditionFieldTitle,
                    this.ConditionOperator,
                    this.ConditionFieldValue);

            this.XmlConditions = String.Format("<Conditions>{0}</Conditions>", conditionXmlBody);
        }
        #endregion
    }
}
