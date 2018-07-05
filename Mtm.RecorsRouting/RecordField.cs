// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RecordField.cs" company="Montrium">
//   MIT License
// </copyright>
// <summary>
//   SPField Updater.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.SharePoint;
    using Microsoft.SharePoint.Taxonomy;

    /// <summary>SPField Updater.</summary>
    public class RecordField : IDisposable
    {
        #region fields
        /// <summary>is Disposed.</summary>
        private bool isDisposed = false;
        #endregion

        #region Constructors
        /// <summary>Initializes a new instance of the <see cref="RecordField"/> class.</summary>
        /// <param name="contextSPField">The context sp field.</param>
        /// <param name="fileSPField">The file SP field.</param>
        public RecordField(SPField contextSPField, SPField fileSPField)
        {
            if (contextSPField == null) throw new ArgumentNullException("contextSPField");
            if (fileSPField == null) throw new ArgumentNullException("fileSPField");

            // contextual values
            this.ContextSPField = contextSPField;
            this.FileSPField = fileSPField;
            
            this.FileSPFieldUser = fileSPField as SPFieldUser;
            this.SPFieldUserIsMultiValue = this.FileSPFieldUser != null && this.FileSPFieldUser.AllowMultipleValues;
            /*if (this.FileSPFieldUser != null) this.SPFieldUserIsMultiValue = this.FileSPFieldUser.AllowMultipleValues;
            else this.SPFieldUserIsMultiValue = false;*/

            this.FileSPFieldLookup = fileSPField as SPFieldLookup;
            this.SPFieldLookupIsMultiValue = this.FileSPFieldLookup != null && this.FileSPFieldLookup.AllowMultipleValues;
            /*if (this.FileSPFieldLookup != null) this.SPFieldLookupIsMultiValue = this.FileSPFieldLookup.AllowMultipleValues;
            else this.SPFieldLookupIsMultiValue = false;*/

            // initialize SPField tests
            this.InitializeSPFieldTests();
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="RecordField"/> class. 
        /// Releases unmanaged resources and performs other cleanup operations before the<see cref="RecordField"/> is reclaimed by garbage collection.</summary>
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class.
        ~RecordField()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets a value indicating whether [SP field user is multi value].</summary>
        public bool SPFieldUserIsMultiValue { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field lookup is multi value].</summary>
        public bool SPFieldLookupIsMultiValue { get; internal set; }

        /// <summary>Gets a value indicating whether IsSPFieldTypeTextOrChoice.</summary>
        public bool SPFieldTypeIsTextOrChoice { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field type is number].</summary>
        public bool SPFieldTypeIsNumber { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field type has same user].</summary>
        public bool SPFieldTypeHasSameUser { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field type is user].</summary>
        public bool SPFieldTypeIsUser { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field type is lookup].</summary>
        public bool SPFieldTypeIsLookup { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field type is lookup or invalid].</summary>
        public bool SPFieldTypeIsLookupOrInvalid { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field type is date time].</summary>
        public bool SPFieldTypeIsDateTime { get; internal set; }

        /// <summary>Gets a value indicating whether [SP field is taxonomy].</summary>
        public bool SPFieldIsTaxonomy { get; internal set; }

        /// <summary>Gets the file SP field lookup.</summary>
        public SPFieldLookup FileSPFieldLookup { get; internal set; }

        /// <summary>Gets the file SP field user.</summary>
        public SPFieldUser FileSPFieldUser { get; internal set; }

        /// <summary>Gets the file SP field.</summary>
        public SPField FileSPField { get; internal set; }

        /// <summary>Gets the context SP field.</summary>
        public SPField ContextSPField { get; internal set; }
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
            string s = "\t\tIsSPFieldTypeTextOrChoice?\t\t" + this.SPFieldTypeIsTextOrChoice + "\n";
            s += "\t\tIsSPFieldTypeNumber?\t\t\t" + this.SPFieldTypeIsNumber + "\n";
            s += "\t\tIsSPFieldTypeSameUser?\t\t\t" + this.SPFieldTypeHasSameUser + "\n";
            s += "\t\tIsSPFieldTypeUser?\t\t\t\t" + this.SPFieldTypeIsUser + "\n";
            s += "\t\tIsSPFieldTypeLookup?\t\t\t" + this.SPFieldTypeIsLookup + "\n";
            s += "\t\tIsSPFieldTypeLookupOrInvalid?\t" + this.SPFieldTypeIsLookupOrInvalid + "\n";
            s += "\t\tIsSPFieldTypeDateTime?\t\t\t" + this.SPFieldTypeIsDateTime + "\n";
            s += "\t\tIsSPFieldTaxonomy?\t\t\t\t" + this.SPFieldIsTaxonomy + "\n";
            s += "\t\tSPFieldUserIsMultiValue?\t\t" + this.SPFieldUserIsMultiValue + "\n";
            s += "\t\tSPFieldLookupIsMultiValue?\t\t" + this.SPFieldLookupIsMultiValue + "\n";
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

        #region PublicMethods
        /// <summary>
        /// Fixes the SP field.
        /// </summary>
        /// <param name="contextSpWeb">
        /// The context sp web.
        /// </param>
        /// <param name="contextSpSite">
        /// The context sp site.
        /// </param>
        /// <param name="fileSpListItem">
        /// The file sp list item.
        /// </param>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        public string FixSPField(SPWeb contextSpWeb, SPSite contextSpSite, SPListItem fileSpListItem, string sourceValue)
        {
            string s = String.Empty;

            if (this.SPFieldTypeIsTextOrChoice)
                s = this.FixSPFieldTypeTextOrChoice(contextSpWeb, fileSpListItem, sourceValue);
            else if (this.SPFieldTypeIsNumber)
                s = this.FixSPFieldTypeNumber(fileSpListItem, sourceValue);
            else if (this.SPFieldTypeHasSameUser)
                s = this.FixSPFieldTypeSameUser(contextSpWeb, fileSpListItem, sourceValue);
            else if (this.SPFieldTypeIsUser)
                s = this.FixSPFieldTypeUser(contextSpWeb, fileSpListItem, sourceValue);
            else if (this.SPFieldTypeIsLookup)
                s = this.FixSPFieldTypeLookup(fileSpListItem, sourceValue);
            else if (this.SPFieldTypeIsDateTime)
                s = this.FixSPFieldTypeDateTime(fileSpListItem, sourceValue);
            else if (this.SPFieldIsTaxonomy)
                s = this.FixSPFieldTaxonomy(contextSpSite, fileSpListItem, sourceValue);
            else
                s = this.FixSPFieldTypeByDefault(fileSpListItem, sourceValue);

            fileSpListItem.UpdateOverwriteVersion();

            return s;
        }
        #endregion

        #region PrivateDisposable
        /// <summary>Releases unmanaged and - optionally - managed resources.</summary>
        /// <param name="isDisposing">The is disposing.</param>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// true to release both managed and unmanaged resources; false to release only unmanaged resources.
        private void Dispose(bool isDisposing)
        {
            // Check if Dispose has been called
            if (!this.isDisposed)
            {
                // dispose managed and unmanaged resources
                if (isDisposing)
                {
                    // managed resources clean
                    this.FileSPFieldLookup = null;
                    this.FileSPFieldUser = null;
                    this.ContextSPField = this.FileSPField = null;
                    this.FileSPFieldLookup = null;
                }

                // unmanaged resources clean
                this.SPFieldUserIsMultiValue = this.SPFieldLookupIsMultiValue = this.SPFieldTypeIsTextOrChoice =
                    this.SPFieldTypeIsNumber = this.SPFieldTypeHasSameUser = this.SPFieldTypeIsUser =
                    this.SPFieldTypeIsLookup = this.SPFieldTypeIsLookupOrInvalid = this.SPFieldTypeIsDateTime =
                    this.SPFieldIsTaxonomy = false;

                // confirm cleaning
                this.isDisposed = true;
            }
        }
        #endregion

        #region PrivateInitializerMethods
        /// <summary>Initializes the SP field tests.</summary>
        private void InitializeSPFieldTests()
        {
            this.IsSPFieldTypeTextOrChoice();
            this.IsSPFieldTypeNumber();
            this.IsSPFieldTypeSameUser();
            this.IsSPFieldTypeUser();
            this.IsSPFieldTypeLookup();
            this.IsSPFieldTypeLookupOrInvalid();
            this.IsSPFieldTypeDateTime();
            this.IsSPFieldTaxonomy();
        }

        /// <summary>Determines whether [is SP field type text or choice].</summary>
        private void IsSPFieldTypeTextOrChoice()
        {
            this.SPFieldTypeIsTextOrChoice = this.ContextSPField.Type == SPFieldType.Text || this.ContextSPField.Type == SPFieldType.Choice;
        }

        /// <summary>Determines whether [is SP field type number].</summary>
        private void IsSPFieldTypeNumber()
        {
            this.SPFieldTypeIsNumber = this.ContextSPField.Type == SPFieldType.Number;
        }

        /// <summary>Determines whether [is SP field type same user] [the specified source sp field].</summary>
        private void IsSPFieldTypeSameUser()
        {
            this.SPFieldTypeHasSameUser = this.ContextSPField.Type == SPFieldType.User && this.FileSPField.Type == SPFieldType.User;
        }

        /// <summary>Determines whether [is SP field type user].</summary>
        private void IsSPFieldTypeUser()
        {
            this.SPFieldTypeIsUser = this.ContextSPField.Type == SPFieldType.User;
        }

        /// <summary>Determines whether [is SP field type lookup].</summary>
        private void IsSPFieldTypeLookup()
        {
            this.SPFieldTypeIsLookup = this.ContextSPField.Type == SPFieldType.Lookup;
        }

        /// <summary>Determines whether [is SP field type lookup or invalid].</summary>
        private void IsSPFieldTypeLookupOrInvalid()
        {
            this.SPFieldTypeIsLookupOrInvalid =
                this.ContextSPField.Type == SPFieldType.Lookup || this.ContextSPField.Type == SPFieldType.Invalid;
        }

        /// <summary>Determines whether [is SP field type date time].</summary>
        private void IsSPFieldTypeDateTime()
        {
            this.SPFieldTypeIsDateTime = this.ContextSPField.Type == SPFieldType.DateTime;
        }

        /// <summary>Determines whether [is SP field taxonomy].</summary>
        private void IsSPFieldTaxonomy()
        {
            this.SPFieldIsTaxonomy = this.ContextSPField is TaxonomyField;
        }
        #endregion

        #region PrivateHelperMethods
        /// <summary>
        /// The Sub string before.
        /// </summary>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <param name="pattern">
        /// The pattern.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        private string SubStringBefore(string sourceValue, string pattern)
        {
            return sourceValue.Split(pattern.ToCharArray())[0];
        }

        /// <summary>
        /// Substrings the after.
        /// </summary>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <param name="pattern">
        /// The pattern.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        private string SubStringAfter(string sourceValue, string pattern)
        {
            if (sourceValue.Split(pattern.ToCharArray()).Length == 1)
                return sourceValue;
            else if (sourceValue.Split(pattern.ToCharArray()).Length == 2)
                return sourceValue.Split(pattern.ToCharArray())[2];
            else
            {
                // create a list and remove the first 2 items
                string result = String.Empty;
                List<string> listString = new List<string>(sourceValue.Split(pattern.ToCharArray()));
                listString.RemoveRange(0, Math.Min(listString.Count, 2));

                foreach (string s in listString) result += s + pattern;

                return result.TrimEnd(pattern.ToCharArray());
            }
        }

        /// <summary>
        /// SPs the field lookup value string.
        /// </summary>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        private string GetSPFieldLookupValues(string sourceValue)
        {
            string formatted = String.Empty;

            foreach (SPFieldLookupValue spFieldLookupValue in new SPFieldLookupValueCollection(sourceValue))
                formatted += ";" + spFieldLookupValue.LookupValue;

            return formatted.TrimStart(';');
        }

        /// <summary>
        /// Gets the SP field lookup values.
        /// </summary>
        /// <param name="contextSpWeb">
        /// The context sp web.
        /// </param>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        private string GetSPFieldLookupValues(SPWeb contextSpWeb, string sourceValue)
        {
            string formatted = String.Empty;

            foreach (SPFieldLookupValue spFieldLookupValue in new SPFieldUserValueCollection(contextSpWeb, sourceValue))
                formatted += ";" + spFieldLookupValue.LookupValue;

            return formatted.TrimStart(';');
        }

        /// <summary>
        /// Gets the SP field lookup value.
        /// </summary>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        private string GetSPFieldLookupValue(string sourceValue) 
        {
            return new SPFieldLookupValue(sourceValue).LookupValue;
        }

        /// <summary>
        /// Gets the SP field lookup value.
        /// </summary>
        /// <param name="sourceValue">
        /// The source value.
        /// </param>
        /// <param name="pattern">
        /// The pattern.
        /// </param>
        /// <returns>
        /// The System.String.
        /// </returns>
        private string GetSPFieldLookupValue(string sourceValue, string pattern)
        {
            return this.SubStringAfter(sourceValue, pattern);
        }

        /// <summary>
        /// Gets the lookup value.
        /// </summary>
        /// <param name="siteId">The site id.</param>
        /// <param name="spFieldLookup">The sp field lookup.</param>
        /// <param name="lookupValue">The lookup value.</param>
        /// <returns>
        /// Returns the SPFieldLookupValue instance of a lookup value.
        /// The ID value will be obtained using SPQuery.
        /// </returns>
        private SPFieldLookupValue GetLookupValue(Guid siteId, SPFieldLookup spFieldLookup, string lookupValue)
        {
            int lookupId;
            string queryFormat =
                @"<Where>
                <Eq>
                    <FieldRef Name='{0}' />
                    <Value Type='Text'>{1}</Value>
                </Eq>
            </Where>";
            using (SPSite spSite = new SPSite(siteId))
            using (SPWeb spWeb = spSite.OpenWeb(spFieldLookup.LookupWebId))
            {
                string queryText = String.Format(CultureInfo.InvariantCulture, queryFormat, spFieldLookup.LookupField, lookupValue);
                SPList lookupSPList = spWeb.Lists.GetList(new Guid(spFieldLookup.LookupList), false);

                SPQuery spQuery = new SPQuery
                    {
                        ViewAttributes = "Scope=\"Recursive\"",
                        ViewFields = "<FieldRef Name='ID'/>",
                        ViewFieldsOnly = true,
                        Query = queryText
                    };

                SPListItemCollection lookupSPListItems = lookupSPList.GetItems(spQuery);

                if (lookupSPListItems.Count > 0)
                {
                    lookupId = Convert.ToInt32(lookupSPListItems[0][SPBuiltInFieldId.ID], CultureInfo.InvariantCulture);
                    return new SPFieldLookupValue(lookupId, lookupValue);
                }
                else
                {
                    // when, as "lookup"-value, this function is passed the actual index ID,
                    // the following section will determine if an item with said ID exists in lookup list
                    if (int.TryParse(lookupValue, out lookupId))
                    {
                        queryFormat = @"<Where>
                                        <Eq>
                                            <FieldRef Name='ID' />
                                            <Value Type='Integer'>{0}</Value>
                                        </Eq>
                                        </Where>";
                        queryText = String.Format(CultureInfo.InvariantCulture, queryFormat, lookupId);
                        spQuery = new SPQuery
                            {
                                ViewAttributes = "Scope=\"Recursive\"",
                                ViewFields =
                                    String.Format(
                                        CultureInfo.InvariantCulture, "<FieldRef Name='{0}'/>", spFieldLookup.LookupField),
                                ViewFieldsOnly = true,
                                Query = queryText
                            };

                        lookupSPListItems = lookupSPList.GetItems(spQuery);

                        if (lookupSPListItems.Count > 0)
                        {
                            lookupValue = lookupSPListItems[0][spFieldLookup.LookupField] as string;
                            if (lookupValue != null)
                                return new SPFieldLookupValue(lookupId, lookupValue);
                        }
                    }
                    // no value found
                    return null;
                }
            }
        }

        /// <summary>
        /// Determines whether [is user SP group name] [the specified context SP web].
        /// </summary>
        /// <param name="contextSPWeb">The context SP web.</param>
        /// <param name="user">The user.</param>
        /// <returns>
        ///   <c>true</c> if [is user SP group name] [the specified context SP web]; otherwise, <c>false</c>.
        /// </returns>
        private bool IsUserSPGroupName(SPWeb contextSPWeb, string user)
        {
            // checking if this string is the name of a SharePoint group
            foreach (SPGroup spGroup in contextSPWeb.SiteGroups)
                if (spGroup.Name.Equals(user))
                    return true;

            return false;
        }

        /// <summary>
        /// Gets the user SP principal.
        /// </summary>
        /// <param name="contextSPWeb">The context SP web.</param>
        /// <param name="user">The user.</param>
        /// <returns>
        /// The SPGroup as SPPrincipal corresponding to the user.
        /// </returns>
        private SPPrincipal GetUserSPPrincipal(SPWeb contextSPWeb, string user)
        {
            // checking if this string is the name of a SharePoint group
            foreach (SPGroup spGroup in contextSPWeb.SiteGroups)
                if (spGroup.Name.Equals(user))
                    return spGroup;

            return null;
        }
        #endregion

        #region PrivateEasyFix
        /// <summary>
        /// Fixes the SP field type number.
        /// </summary>
        /// <param name="fileSPListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>
        /// The fix sp field type number.
        /// </returns>
        private string FixSPFieldTypeNumber(SPListItem fileSPListItem, string sourceValue)
        {
            string s = "\t\tNumber(" + sourceValue + ")\t[";
            // ensure integer values when the target field requires numbers
            sourceValue = this.SubStringAfter(sourceValue, ";#");

            s += sourceValue + "]\n";
            s += "\t\twas:" + fileSPListItem[this.FileSPField.Id] + "\n";
            int num;
            if (int.TryParse(sourceValue, out num)) fileSPListItem[this.FileSPField.Id] = sourceValue;

            s += "\t\tis_:" + fileSPListItem[this.ContextSPField.Id] + "\n";
            return s;
        }

        /// <summary>Fixes the SP field type date time.</summary>
        /// <param name="fileSPListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>The fix sp field type date time.</returns>
        private string FixSPFieldTypeDateTime(SPListItem fileSPListItem, string sourceValue)
        {
            string s = "\t\tDateTime(" + sourceValue + ")\t["
                       + String.Format(DateTimeFormatInfo.InvariantInfo, sourceValue) + "]\n";
            s += "\t\twas:" + fileSPListItem[this.FileSPField.Id] + "\n";

            // ensure that date string values are properly formatted (as an invariant date)
            if (String.IsNullOrEmpty(sourceValue) == false)
                fileSPListItem[this.FileSPField.Id] = String.Format(DateTimeFormatInfo.InvariantInfo, sourceValue);

            s += "\t\tis_:" + fileSPListItem[this.ContextSPField.Id] + "\n";
            return s;
        }

        /// <summary>Fixes the SP field type by default.</summary>
        /// <param name="fileSPListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>The fix sp field type by default.</returns>
        private string FixSPFieldTypeByDefault(SPListItem fileSPListItem, string sourceValue)
        {
            string s = "\t\tDefault(" + sourceValue + ")\n";
            s += "\t\twas:" + fileSPListItem[this.FileSPField.Id] + "\n";

            // can't find an obvious transformation path; copy the value as is
            fileSPListItem[this.FileSPField.Id] = sourceValue;

            s += "\t\tis_:" + fileSPListItem[this.FileSPField.Id] + "\n";
            return s;
        }
        #endregion

        #region PrivateHardFix
        /// <summary>
        /// Fixes the SP field type text or choice.
        /// </summary>
        /// <param name="contextSpWeb">The context sp web.</param>
        /// <param name="fileSpListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>
        /// The fix sp field type text or choice.
        /// </returns>
        private string FixSPFieldTypeTextOrChoice(SPWeb contextSpWeb, SPListItem fileSpListItem, string sourceValue)
        {
            string s = "\t\tTEXT_OR_CHOICE(" + contextSpWeb.Url + "," + sourceValue + ")\n";
            s += "\t\twas:" + fileSpListItem[this.FileSPField.Id] + "\n";

            if (this.SPFieldIsTaxonomy)
            {
                s += "\t\tT_OR_C_IsThisTax? " + this.SPFieldIsTaxonomy + "\n";
                fileSpListItem[this.FileSPField.Id] =
                    this.SubStringBefore(sourceValue,
                                        TaxonomyField.TaxonomyGuidLabelDelimiter.ToString(CultureInfo.InvariantCulture));
                s += "\t\tis_:" + fileSpListItem[this.FileSPField.Id] + "\n";
            }
            else if (this.SPFieldTypeIsLookupOrInvalid)
            {
                s += "\t\tT_OR_C_IsThisLookupOrInvalid? " + this.SPFieldTypeIsLookupOrInvalid + "\n";

                if (this.FileSPFieldLookup == null)
                {
                    s += "\t\tT_OR_C_IsThisLookupOrInvalid_null\n";

                    // multi-valued source fields should be deconstructed properly!
                    fileSpListItem[this.FileSPField.Id] = this.SubStringAfter(sourceValue, ";#");
                    s += "\t\tis_:" + fileSpListItem[this.FileSPField.Id] + "\n";
                }
                else
                {
                    s += "\t\tT_OR_C_IsThisLookupOrInvalid_!null\tIsThisLookupMulti? " + this.SPFieldLookupIsMultiValue + "\n";
                    if (this.SPFieldLookupIsMultiValue)
                        fileSpListItem[this.FileSPField.Id] = this.GetSPFieldLookupValues(sourceValue);
                    else
                        fileSpListItem[this.FileSPField.Id] = this.GetSPFieldLookupValue(sourceValue);

                    s += "\t\tis_:" + fileSpListItem[this.FileSPField.Id] + "\n";
                }
            }
            else if (this.SPFieldTypeIsUser)
            {
                s += "\t\tT_OR_C_IsThisUser? " + this.SPFieldTypeIsUser + "\n";

                // TEST THIS !!!!!!!!
                fileSpListItem[this.FileSPField.Id] = this.GetSPFieldLookupValues(contextSpWeb, sourceValue);
                s += "\t\tis_:" + fileSpListItem[this.FileSPField.Id] + "\n";
            }
            else
            {
                s += "\t\tT_OR_C_else\n";
                fileSpListItem[this.FileSPField.Id] = this.GetSPFieldLookupValue(sourceValue, ";#");
                s += "\t\tis_:" + fileSpListItem[this.FileSPField.Id] + "\n";
            }
            
            return s;
        }

        /// <summary>Fixes the SP field taxonomy.</summary>
        /// <param name="contextSpSite">The context sp site.</param>
        /// <param name="fileSPListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>The fix sp field taxonomy.</returns>
        private string FixSPFieldTaxonomy(SPSite contextSpSite, SPListItem fileSPListItem, string sourceValue)
        {
            string s = "\t\tTaxonomy\n\t\t-" + contextSpSite.Url + ",\n\t\t-" + sourceValue + "\n";
            int lcid = CultureInfo.CurrentCulture.LCID;

            // this is a managed metadata field
            TaxonomyField taxonomyField = this.ContextSPField as TaxonomyField;
            TaxonomySession taxonomySession = new TaxonomySession(contextSpSite);
            TermStore termStore = taxonomySession.TermStores[taxonomyField.SspId];
            TermSet termSet = termStore.GetTermSet(taxonomyField.TermSetId);
            s += "\t\tlcid:\t\t" + lcid + "\n\t\ttaxField:\t" + taxonomyField.Title + "\n";
            s += "\t\ttermstores:\t" + taxonomySession.TermStores.Count + "\n";
            s += "\t\ttermstore:\t" + termStore.Name + "\n\t\ttermset:\t" + termSet.Name + "\n";

            // this is Classification code; to be replaced with the one below!
            Term myTerm = termSet.GetTerms(this.SubStringBefore(sourceValue, "|"), false).FirstOrDefault();

            if (myTerm != null)
            {
                s += "\t\tmyTerm:\t\t" + myTerm.Name + "\n";
                s += "\t\twas:" + fileSPListItem[this.ContextSPField.Id] + "\n";
                string termString = String.Concat(myTerm.GetDefaultLabel(lcid),
                                                    TaxonomyField.TaxonomyGuidLabelDelimiter,
                                                    myTerm.Id);
                int[] ids = TaxonomyField.GetWssIdsOfTerm(
                                                    contextSpSite,
                                                    termStore.Id,
                                                    termSet.Id,
                                                    myTerm.Id,
                                                    true,
                                                    1);
                s += "\t\ttermString:\t" + termString + "\n\t\tids:\t\t" + ids.Length + "\n";

                // set the WssId (TaxonomyHiddenList ID) to -1 so that it is added to the TaxonomyHiddenList
                // fileSPListItem[this.ContextSPField.Id] = (ids.Length == 0) ? "-1;#" + termString : ids[0] + ";#" + termString;
                if (ids.Length == 0)
                    termString = "-1;#" + termString;
                else
                    termString = ids[0] + ";#" + termString;

                fileSPListItem[this.ContextSPField.Id] = termString;
            }

            s += "\t\tis_:" + fileSPListItem[this.ContextSPField.Id] + "\n";
            return s;
        }

        /// <summary>
        /// Fixes the SP field type lookup.
        /// </summary>
        /// <param name="fileSpListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>Fixes the SPField type Lookup.</returns>
        private string FixSPFieldTypeLookup(SPListItem fileSpListItem, string sourceValue)
        {
            string s = "\t\tLookup(" + sourceValue + ")\tIsThisLookup? " + this.SPFieldTypeIsLookup + "\n";
            s += "\t\twas:" + fileSpListItem[this.ContextSPField.Id] + "\n";

            // process lookup columns (including multi-valued ones)
            List<string> lookupValueList = new List<string>();

            if (this.SPFieldTypeIsLookup)
            {
                s += "\t\tIsThisLookupupMulti? " + this.SPFieldLookupIsMultiValue + "\n";

                if (this.SPFieldLookupIsMultiValue)
                    foreach (SPFieldLookupValue spFieldLookupValue in (fileSpListItem[this.FileSPField.Id] as SPFieldLookupValueCollection))
                        lookupValueList.Add(spFieldLookupValue.LookupValue);
                else
                    lookupValueList.Add(new SPFieldLookupValue(sourceValue).LookupValue);
            }
            else
            {
                string[] strList = sourceValue.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                foreach (string str in strList)
                    lookupValueList.Add(str);
                s += "\t\tLookupElse ListCount:\t" + lookupValueList.Count + "\n";
            }

            // construct an SPFieldLookup object in the image of our field
            SPFieldLookup spFieldLookup = this.ContextSPField as SPFieldLookup;
            s += "\t\tspfield as spfield lookup multi? " + spFieldLookup.AllowMultipleValues + "\n";
            s += "\t\tlookupvaluelistcount:\t" + lookupValueList.Count + "\n";

            if (spFieldLookup.AllowMultipleValues)
            {
                SPFieldLookupValueCollection spFieldLookupValues = new SPFieldLookupValueCollection();

                foreach (string value in lookupValueList)
                {
                    // obtain the lookup value and its corresponding ID
                    SPFieldLookupValue lookupValue =
                        GetLookupValue(fileSpListItem.Web.Site.ID, spFieldLookup, value);

                    // lookup value found in the lookup list
                    if (lookupValue != null) spFieldLookupValues.Add(lookupValue);
                }

                fileSpListItem[this.ContextSPField.Id] = spFieldLookupValues;
                s += "\t\tis_:" + spFieldLookupValues + "\n";
            }
            else if (lookupValueList.Count > 0)
            {
                // empty column values are "copied" as well
                if (String.IsNullOrEmpty(lookupValueList[0]))
                {
                    fileSpListItem[this.ContextSPField.Id] = String.Empty;
                    // continue;
                }

                // obtain the lookup value and its corresponding ID
                SPFieldLookupValue spFieldLookupValue =
                    GetLookupValue(fileSpListItem.Web.Site.ID, spFieldLookup, lookupValueList[0]);

                // lookup value found in the lookup list
                if (spFieldLookupValue != null)
                    fileSpListItem[this.ContextSPField.Id] = spFieldLookupValue.ToString();
                s += "\t\tis_:" + fileSpListItem[this.ContextSPField.Id] + "\n";
            }

            return s;
        }

        /// <summary>
        /// Fixes the SP field type same user.
        /// </summary>
        /// <param name="contextSPWeb">The context sp web.</param>
        /// <param name="fileSPListItem">The file SP list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>
        /// The fix to sp field having the same username.
        /// </returns>
        private string FixSPFieldTypeSameUser(SPWeb contextSPWeb, SPListItem fileSPListItem, string sourceValue)
        {
            string s = "\t\tSameUser(" + contextSPWeb.Url + "," + sourceValue + ")\n";
            s += "\t\twas:" + fileSPListItem[this.FileSPField.Id] + "\n";

            // process "Person or Group" fields (including multi-valued ones)
            SPFieldUserValueCollection sourceUserValues = new SPFieldUserValueCollection(contextSPWeb, sourceValue);
            SPFieldUserValueCollection targetUserValues = new SPFieldUserValueCollection();
            s += "\t\tsourceUserValues:" + sourceUserValues.Count + "\n";

            foreach (SPFieldUserValue spFieldUserValue in sourceUserValues)
            {
                string logon = spFieldUserValue.LookupValue;
                SPPrincipal spPrincipal = null;

                // checking if this string is the name of a SharePoint group
                spPrincipal = this.GetUserSPPrincipal(contextSPWeb, logon);

                if (spPrincipal == null)
                {
                    // the value is not a group; try to set it as a user then.
                    try
                    {
                        spPrincipal = contextSPWeb.EnsureUser(logon);
                    }
                    catch
                    {
                    }
                }

                // add to user collection
                if (spPrincipal != null)
                    targetUserValues.Add(new SPFieldUserValue(contextSPWeb, spPrincipal.ID, spPrincipal.Name));
            }
            
            // set multi-user fields to new user collection 
            if (this.SPFieldUserIsMultiValue)
                fileSPListItem[this.FileSPField.Id] = targetUserValues;

            // set single-user fields to just one user value
            else
                fileSPListItem[this.FileSPField.Id] = targetUserValues[0];

            s += "\t\tis_:" + fileSPListItem[this.FileSPField.Id] + "\n";
            return s;
        }
        
        /// <summary>Fixes the SP field type user.</summary>
        /// <param name="contextSPWeb">The context sp web.</param>
        /// <param name="fileSPListItem">The file sp list item.</param>
        /// <param name="sourceValue">The source value.</param>
        /// <returns>The fix sp field type user.</returns>
        private string FixSPFieldTypeUser(SPWeb contextSPWeb, SPListItem fileSPListItem, string sourceValue)
        {
            string s = "\t\tUser(" + contextSPWeb.Url + "," + sourceValue + "\n";
            s += "\t\twas:" + fileSPListItem[this.FileSPField.Id] + "\n";

            // transform non-user values into user type
            string[] userNames = sourceValue.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            s += "\t\tuserNamescount:\t" + userNames.Length + "\n";

            // create list of user names parsing array of names
            SPFieldUserValueCollection spFieldUserValues = new SPFieldUserValueCollection();

            // update SPPrincipal using user info
            foreach (string userName in userNames)
            {
                string user = userName.Trim();
                s += "\t\tuser:\t" + user + "\n";
                SPPrincipal spPrincipal = null;

                // checking if this string is the name of a SharePoint group, add to user collection if (spPrincipal != null)
                if (this.IsUserSPGroupName(contextSPWeb, user))
                {
                    spPrincipal = this.GetUserSPPrincipal(contextSPWeb, user);
                    spFieldUserValues.Add(new SPFieldUserValue(contextSPWeb, spPrincipal.ID, spPrincipal.Name));
                }

                // the value is not a group; try to set it as a user then.
                // if (spPrincipal == null)
                else
                {
                    try
                    {
                        spPrincipal = contextSPWeb.EnsureUser(user);
                    }
                    catch
                    {
                    }
                }
                s += "\t\tspPrncipal:\t" + spPrincipal.Name + "\n";
            }

            // set multi-user fields to new user collection 
            if (this.SPFieldUserIsMultiValue)
                fileSPListItem[this.FileSPField.Id] = spFieldUserValues;

            // set single-user fields to just one user value
            else
                fileSPListItem[this.FileSPField.Id] = spFieldUserValues[0];

            s += "\t\tis_:" + fileSPListItem[this.FileSPField.Id] + "\n";
            return s;
        }
        #endregion
    }
}