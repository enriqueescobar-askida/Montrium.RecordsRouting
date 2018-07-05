// --------------------------------------------------------------------------------------------------------------------
// <copyright file="RoutingRulesManager.cs" company="Montrium">
//   MIT License
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


    /// <summary>ContentOrganizer Manager.</summary>
    public class RoutingRulesManager : IDisposable
    {
        #region fields
        /// <summary>Is Disposed.</summary>
        private bool IsDisposed = false;

        /// <summary>Content Organizer Rules Title.</summary>
        private const string ContentOrganizerRulesTitle = "Content Organizer Rules";

        /// <summary>The routing rules title.</summary>
        private const string RoutingRulesTitle = "RoutingRules";
        #endregion

        #region Constructor
        /// <summary>Initializes a new instance of the <see cref="RoutingRulesManager"/> class.</summary>
        /// <param name="url">The URL.</param>
        public RoutingRulesManager(string url)
        {
            if (String.IsNullOrEmpty(url)) throw new ArgumentNullException(url, "URL NULL ARG0");

            url = Path.Combine(url, RoutingRulesTitle);

            using (SPSite spSite = new SPSite(url))
            using (SPWeb spWeb = spSite.OpenWeb())
            {
                this.RoutingRules = this.ScanForCustomOrganizerRules(spWeb.Lists);
            }
        }
        #endregion

        #region DestructorDisposable
        /// <summary>Finalizes an instance of the <see cref="RoutingRulesManager"/> class. Releases unmanaged resources and performs other cleanup operations before the<see cref="RoutingRulesManager"/> is reclaimed by garbage collection.</summary>
        /// Releases unmanaged resources and performs other cleanup operations before the
        /// is reclaimed by garbage collection.
        /// This destructor will run only if the Dispose method does not get called.
        /// It gives your base class the opportunity to finalize.
        /// Do not provide destructors in types derived from this class
        ~RoutingRulesManager()
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of readability and maintainability.
            this.Dispose(false);
        }
        #endregion

        #region AttributesOrProperties
        /// <summary>Gets Rules.</summary>
        public SPListItemCollection RoutingRules { get; internal set; }
        #endregion

        #region PublicMethods
        /// <summary>
        /// Gets the content type boolean.
        /// </summary>
        /// <param name="recordDocument">
        /// The record document.
        /// </param>
        /// <param name="checkParent">
        /// if set to <c>true</c> [check parent].
        /// </param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        public bool GetContentTypeBoolean(RecordDocument recordDocument, bool checkParent)
        {
            if (!checkParent) return this.GetContentTypeBoolean(recordDocument);
            else
            {
                bool boo = false;
                const string SubContentType = "Submission Content Type";

                foreach (SPListItem routingRule in this.RoutingRules)
                {
                    SPField spField = routingRule.Fields[SubContentType];
                    if (routingRule[spField.Id].Equals(recordDocument.ParentContentType.Name.Trim()))
                        boo = true;
                }

                return boo || this.GetContentTypeBoolean(recordDocument);
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
                const string SubContentType = "Submission Content Type";

                for (int i = 0; i < this.RoutingRules.Count; i++)
                {
                    SPListItem routingRule = this.RoutingRules[i];
                    SPField spField = routingRule.Fields[SubContentType];
                    if (routingRule[spField.Id].Equals(recordDocument.ParentContentType.Name.Trim()))
                        intList.Add(i);
                }

                return intList;
            }
        }

        /// <summary>
        /// Adds the new rule.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <param name="name">The name.</param>
        /// <param name="description">The description.</param>
        /// <param name="libraryName">The library.</param>
        /// <param name="relativeFolderUrl">The relative folder URL.</param>
        /// <param name="contentTypeName">Name of the content type.</param>
        /// <param name="conditionFieldId">The condition field id.</param>
        /// <param name="conditionFieldInternalName">Name of the condition field internal.</param>
        /// <param name="conditionFieldTitle">The condition field title.</param>
        /// <param name="conditionOperator">The condition operator.</param>
        /// <param name="conditionFieldValue">The condition field value.</param>
        public void AddNewRule(
            string url,
            string name,
            string description,
            string libraryName,
            string relativeFolderUrl,
            string contentTypeName,
            string conditionFieldId,
            string conditionFieldInternalName,
            string conditionFieldTitle,
            string conditionOperator,
            string conditionFieldValue)
        {
            List<EcmDocumentRouterRule> ruleList = new List<EcmDocumentRouterRule>();

            // Build the conditionSettings XML from the constants above.
            string conditionXml = String.Format(
                @"<Condition Column=""{0}|{1}|{2}"" Operator=""{3}"" Value=""{4}"" />",
                conditionFieldId,
                conditionFieldInternalName,
                conditionFieldTitle,
                conditionOperator,
                conditionFieldValue);
            string conditionsXml = String.Format("<Conditions>{0}</Conditions>", conditionXml);

            using (SPSite spSite = new SPSite(url))
            using (SPWeb spWeb = spSite.OpenWeb())
            {
                EcmDocumentRoutingWeb edrw = new EcmDocumentRoutingWeb(spWeb);
                SPContentType ruleContentType = spWeb.ContentTypes[contentTypeName];
                SPList ruleLibrary = spWeb.Lists[libraryName];

                if (ruleLibrary.ContentTypes.BestMatch(ruleContentType.Id) == null)
                    throw new ArgumentException(String.Format("Ensure that the library {0} contains content type {1} before creating the rule", libraryName, contentTypeName));

                // Create a blank rule. Configure the rule..
                EcmDocumentRouterRule edrr = new EcmDocumentRouterRule(spWeb)
                    {
                        Name = name,
                        Description = description,
                        ContentTypeString = ruleContentType.Name,
                        RouteToExternalLocation = false,
                        Priority = "5",
                        TargetPath = spWeb.GetFolder(url + relativeFolderUrl).ServerRelativeUrl,
                        ConditionsString = conditionsXml
                    };

                // Update the rule and commit changes.
                edrr.Update();
                ruleList.Add(edrr);
            }
        }

        /// <summary>
        /// Prints to log file.
        /// </summary>
        /// <param name="logPath">The log path.</param>
        public void PrintToLogFile(string logPath)
        {
            // if (String.IsNullOrEmpty(logPath)) throw new ArgumentNullException("logPath");

            using (StreamWriter sw = new StreamWriter(logPath.Replace(".txt", "Rules.txt"), false, Encoding.UTF8))
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
            string s = "Url:\t\t" + this.RoutingRules.Count + "\n|\n";

            foreach (SPListItem routingRule in this.RoutingRules)
            {
                s += routingRule.ContentType.Name.PadRight(15, '=') + routingRule.DisplayName + "\n";
                s += "Name:\t\t" + routingRule.Name + "\n";
                s += "Title:\t\t" + routingRule.Title + "\n";
                s += "WebUrl:\t\t" + routingRule.Web.Url + "\n";
                s += "Url:\t\t" + routingRule.Url + "\n";
                s += "Folder:\t\t" + routingRule.Folder + "\n";
                s += "SortType:\t" + routingRule.SortType + "\n";
                s += "CpSource:\t" + routingRule.CopySource + "\n";
                s += "CopyTo:\t\t" + routingRule.Web.Url + "/" + routingRule["Target Library"] + "\n";
                s += "EBasePerm:\t" + routingRule.EffectiveBasePermissions + "\n";
                s += "HavePerm:\t" + routingRule.DoesUserHavePermissions(routingRule.EffectiveBasePermissions) + "\n";
                s += "FileSysObjT:" + routingRule.FileSystemObjectType + "\n";
                s += "HasUniqRolA:" + routingRule.HasUniqueRoleAssignments + "\n";
                s += "Level:\t\t" + routingRule.Level + "\n";
                s += "MReqFields:\t" + routingRule.MissingRequiredFields + "\n";
                s += "SubCType:\t" + routingRule["Submission Content Type"] + "\n";
                s += "TargetLib:\t" + routingRule["Target Library"] + "\n";
                s += "TargetDir:\t" + routingRule["Target Folder"] + "\n";
                s += "TargetPat:\t" + routingRule["Target Path"] + "\n";
                s += "CustoRout:\t" + routingRule["Custom Router"] + "\n";
                s += "IsActivat:\t" + routingRule["Active"] + "\n";
                s += "...............\n";
            }

            s += "\n|\n";
            return s;
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
        /// <summary>Dispose Finally.</summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// true to release both managed and unmanaged resources; false to release only unmanaged resources.
        /// <param name="isDisposing">The is disposing.</param>
        private void Dispose(bool isDisposing)
        {
            // Check if Dispose has been called
            if (!this.IsDisposed)
            {
                // dispose managed and unmanaged resources
                if (isDisposing)
                {
                    // managed resources clean
                    this.RoutingRules = null;
                }

                // unmanaged resources clean

                // confirm cleaning
                this.IsDisposed = true;
            }
        }
        #endregion

        #region PrivateMethods
        /// <summary>
        /// Gets the content type boolean.
        /// </summary>
        /// <param name="recordDocument">The record document.</param>
        /// <returns>
        /// The System.Boolean.
        /// </returns>
        private bool GetContentTypeBoolean(RecordDocument recordDocument)
        {
            const string SubContentType = "Submission Content Type";

            foreach (SPListItem routingRule in this.RoutingRules)
            {
                SPField spField = routingRule.Fields[SubContentType];
                if (routingRule[spField.Id].Equals(recordDocument.ContentType.Name.Trim())) return true;
            }

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
            const string SubContentType = "Submission Content Type";

            for (int i = 0; i < this.RoutingRules.Count; i++)
            {
                SPListItem routingRule = this.RoutingRules[i];
                SPField spField = routingRule.Fields[SubContentType];
                if (routingRule[spField.Id].Equals(recordDocument.ContentType.Name.Trim()))
                    intList.Add(i);
            }

            return intList;
        }

        /// <summary>
        /// Scans for custom organizer rules.
        /// </summary>
        /// <param name="spListCollection">The sp list collection.</param>
        /// <returns>SPListCollection of custom organizer rules.</returns>
        private SPListItemCollection ScanForCustomOrganizerRules(SPListCollection spListCollection)
        {
            foreach (SPList spList in spListCollection)
                if ((!(spList is SPDocumentLibrary) && !(spList is SPPictureLibrary))
                    && spList.Title.Contains(ContentOrganizerRulesTitle))
                    return spList.Items;

            return null;
        }
        #endregion
    }
}
