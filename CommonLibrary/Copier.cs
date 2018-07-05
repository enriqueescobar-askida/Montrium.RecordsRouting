// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Copier.cs" company="MTM">
//   MIT License
// </copyright>
// <summary>
//   Class Copier.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;

    using Microsoft.SharePoint;
    using Microsoft.SharePoint.Taxonomy;
    using Microsoft.SharePoint.Utilities;

    /// <summary>Class Copier.</summary>
    public class Copier
    {
        /// <summary>Copy FileMetadata.</summary>
        /// <param name="fromFileUrl">The from file url.</param>
        /// <param name="toFileUrl">The to file url.</param>
        /// <param name="omitFields">The omit fields.</param>
        /// <param name="changeVersion">The change version.</param>
        /// <returns>The copy file metadata.</returns>
        public string CopyFileMetadata(string fromFileUrl, string toFileUrl, string omitFields, bool changeVersion)
        {
            string response = string.Empty;

            SPSite iSite = null;
            SPWeb iWeb = null;
            SPSite oSite = null;
            SPWeb oWeb = null;

            try
            {
                iSite = new SPSite(fromFileUrl);
                iWeb = iSite.RootWeb;
                oSite = new SPSite(toFileUrl);
                oWeb = oSite.RootWeb;

                // try and obtain valid instances of the SPFile objects at the specified URLs
                SPFile iFile = this.GetFileObject(iSite, fromFileUrl);
                SPFile oFile = this.GetFileObject(oSite, toFileUrl);

                if (iFile == null) throw new FileNotFoundException("unable to find file " + fromFileUrl);
                if (oFile == null) throw new FileNotFoundException("unable to find file " + toFileUrl);

                SPListItem sourceItem = iFile.Item;
                SPListItem targetItem = oFile.Item;

                if (sourceItem == null)
                {
                    MTMLogger.High(this.stERR + "CopyFileMetadata :: File " + fromFileUrl + " has no corresponding item.");
                    return this.stERR + "CopyFileMetadata :: File " + fromFileUrl + " has no corresponding item.";
                }
                if (targetItem == null)
                {
                    MTMLogger.High(this.stERR + "CopyFileMetadata :: File " + toFileUrl + " has no corresponding item.");
                    return this.stERR + "CopyFileMetadata :: File " + toFileUrl + " has no corresponding item.";
                }

                string[] blackList = omitFields.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                // trim whitespace from both ends of the elements
                for (int i = 0; i < blackList.Length; i++)
                    blackList[i] = blackList[i].Trim();

                // disable the security validation (temporarily)
                oWeb.AllowUnsafeUpdates = true;

                SPField sourceField = null;
                foreach (SPField targetField in targetItem.Fields)
                {
                    if (((IList)blackList).Contains(targetField.Title)
                        || this.FieldIsCopyTarget(targetField, sourceItem, ref sourceField) == false)
                    {
                        continue;
                    }

                    try
                    {
                        string sourceValue = Convert.ToString(sourceItem[sourceField.Id]);

                        // TODO: populating the choice-type column has to be separate!
                        if (this.IsTargetSPFieldTypeTextOrChoice(targetField))
                            this.MatchSPFieldTextOrChoice(targetItem, targetField, iWeb, sourceField, sourceValue);
                        else if (this.IsTargetSPFieldTypeNumber(targetField))
                            this.MatchSPFieldTypeNumber(targetItem, targetField, sourceValue);
                        else if (this.IsSPFieldTypeUserIdem(targetField, sourceField))
                            this.MatchSPFieldTypeUserIdem(iWeb, oWeb, targetItem, targetField, sourceValue);
                        else if (this.IsTargetSPFieldTypeUser(targetField))
                            this.MatchSPFieldTypeUser(oWeb, targetItem, targetField, sourceValue);
                        else if (this.IsTargetSPFieldTypeLookup(targetField))
                            this.MatchSPFieldTypeLookup(targetItem, targetField, sourceItem, sourceField, sourceValue);
                        else if (this.IsTargetSPFieldTypeDateTime(targetField))
                            this.MatchSPFieldTypeDateTime(targetItem, targetField, sourceValue);
                        else if (this.IsTargetFieldTaxonomyField(targetField))
                            this.MatchTaxonomyField(targetItem, targetField, sourceValue);
                        else
                            this.MatchSPFieldDefaultId(targetItem, targetField, sourceItem, sourceField);
                    }
                    catch (Exception ex)
                    {
                        // catch the field update exception and move on
                        response += "Warning: [" + targetField.Title + "] " + ex.Message + '\n';
                    }
                }

                // Set "Document Source" metadata to the source file's URL
                if (targetItem.Fields.ContainsField("Document Source") && !((IList)blackList).Contains("Document Source"))
                {
                    SPField field = targetItem.Fields["Document Source"];

                    // this is a hyperlink-type column
                    if (field.Type == SPFieldType.URL)
                    {
                        SPFieldUrlValue newLink = new SPFieldUrlValue();
                        newLink.Description = "Link to: " + sourceItem.Name;
                        newLink.Url = SPEncode.UrlEncodeAsUrl(iWeb.Url + '/' + sourceItem.Url);
                        targetItem["Document Source"] = newLink;
                    }
                    else
                    {
                        // or just a "single line of text" column
                        targetItem["Document Source"] = SPEncode.UrlEncodeAsUrl(iWeb.Url + '/' + sourceItem.Url);
                    }
                }

                // apply metadata update without triggering any events
                using (new DisabledItemEventsScope())
                {
                    // don't create a new version; the item is updated regardless of checkout status
                    if (changeVersion == false)
                    {
                        try
                        {
                            targetItem.SystemUpdate(false);
                        }
                        catch (Exception ex)
                        {
                            MTMLogger.High(this.stERR + "CopyFileMetadata :: " + ex.Message);
                            response = this.stERR + "CopyFileMetadata :: " + ex.Message;
                        }
                    }

                    // follow proper checkout-update-checkin procedure
                    else
                    {
                        // check if the document is checked in
                        // (depends on the security settings of the document library)
                        if (oFile.CheckOutType == SPFile.SPCheckOutType.None)
                        {
                            try
                            {
                                oFile.CheckOut();
                                // Update the item
                                targetItem.Update();
                                oFile.CheckIn(string.Empty);
                            }
                            catch (Exception ex)
                            {
                                MTMLogger.High(stERR + "CopyFileMetadata :: " + ex.Message);
                                response = stERR + "CopyFileMetadata :: " + ex.Message;
                            }
                        }
                        else
                        {
                            try
                            {
                                // Update the item
                                targetItem.Update();
                            }
                            catch (Exception ex)
                            {
                                MTMLogger.High(stERR + "CopyFileMetadata :: " + ex.Message);
                                response = stERR + "CopyFileMetadata :: " + ex.Message;
                            }
                        }
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                MTMLogger.High(stERR + "CopyFileMetadata :: " + ex.Message);
                response = stERR + "CopyFileMetadata :: " + ex.Message;
            }
            catch (Exception ex)
            {
                MTMLogger.High(stERR + "CopyFileMetadata :: " + ex.Message);
                response = stERR + "CopyFileMetadata :: " + ex.Message;
            }
            finally
            {
                if (iWeb != null) iWeb.Dispose();
                if (iSite != null) iSite.Dispose();
                if (oWeb != null) oWeb.Dispose();
                if (oSite != null) oSite.Dispose();
            }

            return response;
        }

        #region PrivateInitMethods
        private bool IsTargetSPFieldTypeTextOrChoice(SPField targetSPField)
        {
            return targetSPField.Type == SPFieldType.Text || targetSPField.Type == SPFieldType.Choice;
        }
        private bool IsTargetSPFieldTypeNumber(SPField targetSPField)
        {
            return targetSPField.Type == SPFieldType.Number;
        }
        private bool IsSPFieldTypeUserIdem(SPField targetSPField, SPField sourceSPField)
        {
            return (targetSPField.Type == SPFieldType.User) && (sourceSPField.Type == SPFieldType.User);
        }
        private bool IsTargetSPFieldTypeUser(SPField targetSPField)
        {
            return targetSPField.Type == SPFieldType.User;
        }
        private bool IsTargetSPFieldTypeLookup(SPField targetSPField)
        {
            return targetSPField.Type == SPFieldType.Lookup;
        }
        private bool IsTargetSPFieldTypeDateTime(SPField targetSPField)
        {
            return targetSPField.Type == SPFieldType.DateTime;
        }
        private bool IsTargetFieldTaxonomyField(SPField targetSPField)
        {
            return targetSPField is TaxonomyField;
        }
        #endregion

        #region PrivateMatchingMethods
        private void MatchSPFieldTextOrChoice(SPListItem targetSPListItem, SPField targetSPField,
            SPWeb sourceSPWeb, SPField sourceSPField, string sourceValue)
        {
            // get just the term label
            if (sourceSPField is TaxonomyField)
                targetSPListItem[targetSPField.Id] =
                    this.SubstringBefore(sourceValue,
                    TaxonomyField.TaxonomyGuidLabelDelimiter.ToString(CultureInfo.InvariantCulture));

            // for lookup and custom fields, convert RAW values into plain text
            else if (sourceSPField.Type == SPFieldType.Lookup || sourceSPField.Type == SPFieldType.Invalid)
            {
                SPFieldLookup spFieldLookupField = sourceSPField as SPFieldLookup;
                if (spFieldLookupField == null)
                {
                    // TODO: multi-valued source fields should be deconstructed properly!
                    targetSPListItem[targetSPField.Id] = this.SubstringAfter(sourceValue, ";#");
                }
                else
                {
                    if (spFieldLookupField.AllowMultipleValues)
                    {
                        string formatted = string.Empty;
                        SPFieldLookupValueCollection spFieldLookupValues =
                            new SPFieldLookupValueCollection(sourceValue);
                        foreach (SPFieldLookupValue spFieldLookupValue in spFieldLookupValues)
                            formatted = formatted + ";" + spFieldLookupValue.LookupValue;

                        targetSPListItem[targetSPField.Id] = formatted.TrimStart(';');
                    }
                    else
                    {
                        SPFieldLookupValue spFieldLookupValue = new SPFieldLookupValue(sourceValue);
                        targetSPListItem[targetSPField.Id] = spFieldLookupValue.LookupValue;
                    }
                }
            }
            else if (sourceSPField.Type == SPFieldType.User)
            {
                // TODO: TEST THIS !!!!!!!!
                string formatted = string.Empty;
                SPFieldUserValueCollection spFieldUserValues = new SPFieldUserValueCollection(sourceSPWeb, sourceValue);
                foreach (SPFieldLookupValue spFieldLookupValue in spFieldUserValues)
                    formatted = formatted + ";" + spFieldLookupValue.LookupValue;

                targetSPListItem[targetSPField.Id] = formatted.TrimStart(';');
            }
            else targetSPListItem[targetSPField.Id] = this.SubstringAfter(sourceValue, ";#");
        }

        private void MatchSPFieldTypeNumber(SPListItem targetSPListItem, SPField targetSPField, string sourceValue)
        {
            // ensure integer values when the target field requires numbers
            sourceValue = this.SubstringAfter(sourceValue, ";#");

            int num;
            if (int.TryParse(sourceValue, out num)) targetSPListItem[targetSPField.Id] = sourceValue;
        }

        private void MatchSPFieldTypeUserIdem(SPWeb fileSPWeb, SPWeb contextSPWeb, SPListItem targetSPListItem,
            SPField targetSPField, string sourceValue)
        {
            // process "Person or Group" fields (including multi-valued ones)
            SPFieldUserValueCollection sourceUsers = new SPFieldUserValueCollection(fileSPWeb, sourceValue);

            if (fileSPWeb.ID == contextSPWeb.ID)
            {
                // set multi-user fields to user collection 
                if ((targetSPField as SPFieldUser).AllowMultipleValues)
                    targetSPListItem[targetSPField.Id] = sourceUsers;

                // set single-user fields to just one user value
                else
                    targetSPListItem[targetSPField.Id] = sourceUsers[0];
            }
            else
            {
                SPFieldUserValueCollection targetUsers = new SPFieldUserValueCollection();

                foreach (SPFieldUserValue userValue in sourceUsers)
                {
                    string logon = userValue.LookupValue;
                    SPPrincipal principal = null;

                    // checking if this string is the name of a SharePoint group
                    foreach (SPGroup oGroup in contextSPWeb.SiteGroups)
                    {
                        if (oGroup.Name == logon)
                        {
                            principal = oGroup;
                            break;
                        }
                    }
                    if (principal == null)
                    {
                        // the value is not a group; try to set it as a user then.
                        try
                        {
                            principal = contextSPWeb.EnsureUser(logon);
                        }
                        catch
                        {
                        }
                    }
                    // add to user collection
                    if (principal != null)
                        targetUsers.Add(new SPFieldUserValue(contextSPWeb, principal.ID, principal.Name));
                }

                // set multi-user fields to new user collection 
                if ((targetSPField as SPFieldUser).AllowMultipleValues)
                    targetSPListItem[targetSPField.Id] = targetUsers;

                // set single-user fields to just one user value
                else
                    targetSPListItem[targetSPField.Id] = targetUsers[0];
            }
        }

        private void MatchSPFieldTypeUser(SPWeb contextSPWeb, SPListItem targetSPListItem, SPField targetSPField,
            string sourceValue)
        {
            // transform non-user values into user type
            string[] nameArr = sourceValue.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            // create list of user names
            SPFieldUserValueCollection spFieldUserValueCollection = new SPFieldUserValueCollection();

            foreach (string name in nameArr)
            {
                SPPrincipal spPrincipal = null;

                // checking if this string is the name of a SharePoint group
                foreach (SPGroup spGroup in contextSPWeb.SiteGroups)
                {
                    if (spGroup.Name == name.Trim())
                    {
                        spPrincipal = spGroup;
                        break;
                    }
                }

                // the value is not a group; try to set it as a user then.
                if (spPrincipal == null)
                {
                    try
                    {
                        spPrincipal = contextSPWeb.EnsureUser(name.Trim());
                    }
                    catch
                    {
                    }
                }

                // add to user collection
                if (spPrincipal != null)
                    spFieldUserValueCollection.Add(
                        new SPFieldUserValue(contextSPWeb, spPrincipal.ID, spPrincipal.Name));
            }

            // set multi-user fields to new user collection 
            if ((targetSPField as SPFieldUser).AllowMultipleValues)
                targetSPListItem[targetSPField.Id] = spFieldUserValueCollection;

            // set single-user fields to just one user value
            else
                targetSPListItem[targetSPField.Id] = spFieldUserValueCollection[0];
        }

        private void MatchSPFieldTypeLookup(SPListItem targetSPListItem, SPField targetSPField,
            SPListItem sourceSPListItem, SPField sourceSPField,  string sourceValue)
        {
            // process lookup columns (including multi-valued ones)
            List<string> valueList = new List<string>();

            if (sourceSPField.Type == SPFieldType.Lookup)
            {
                if ((sourceSPField as SPFieldLookup).AllowMultipleValues)
                {
                    SPFieldLookupValueCollection spFieldLookupValues =
                        sourceSPListItem[sourceSPField.Id] as SPFieldLookupValueCollection;

                    foreach (SPFieldLookupValue value in spFieldLookupValues)
                        valueList.Add(value.LookupValue);
                }
                else
                {
                    SPFieldLookupValue spFieldLookupValue = new SPFieldLookupValue(sourceValue);
                    valueList.Add(spFieldLookupValue.LookupValue);
                }
            }
            else
            {
                string[] strList = sourceValue.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string str in strList)
                    valueList.Add(str);
            }

            // construct an SPFieldLookup object in the image of our field
            SPFieldLookup spFieldLookup = targetSPField as SPFieldLookup;

            if (spFieldLookup.AllowMultipleValues)
            {
                SPFieldLookupValueCollection spFieldLookupValues = new SPFieldLookupValueCollection();

                foreach (string value in valueList)
                {
                    // obtain the lookup value and its corresponding ID
                    SPFieldLookupValue lookupValue =
                        GetLookupValue(targetSPListItem.Web.Site.ID, spFieldLookup, value);

                    // lookup value found in the lookup list
                    if (lookupValue != null)
                        spFieldLookupValues.Add(lookupValue);
                }

                targetSPListItem[targetSPField.Id] = spFieldLookupValues;
            }
            else if (valueList.Count > 0)
                {
                    // empty column values are "copied" as well
                    if (string.IsNullOrEmpty(valueList[0]))
                    {
                        targetSPListItem[targetSPField.Id] = string.Empty;
                        // continue;
                    }

                    // obtain the lookup value and its corresponding ID
                    SPFieldLookupValue spFieldLookupValue =
                        GetLookupValue(targetSPListItem.Web.Site.ID, spFieldLookup, valueList[0]);

                    // lookup value found in the lookup list
                    if (spFieldLookupValue != null)
                        targetSPListItem[targetSPField.Id] = spFieldLookupValue.ToString();
                }
        }

        private void MatchSPFieldTypeDateTime(SPListItem targetSPListItem, SPField targetSPField, string sourceValue)
        {
            // ensure that date string values are properly formatted (as an invariant date)
            if (string.IsNullOrEmpty(sourceValue) == false)
                targetSPListItem[targetSPField.Id] = string.Format(DateTimeFormatInfo.InvariantInfo, sourceValue);
        }

        private void MatchTaxonomyField(SPListItem targetSPListItem, SPField targetSPField, string sourceValue)
        {
            // this is a managed metadata field
            TaxonomyField managedField = targetSPField as TaxonomyField;
            TaxonomySession session = new TaxonomySession(targetSPListItem.Web.Site);
            TermStore termStore = session.TermStores[managedField.SspId];
            TermSet termSet = termStore.GetTermSet(managedField.TermSetId);
            int lcid = CultureInfo.CurrentCulture.LCID;

            // TODO: this is Classification code; to be replaced with the one below!
            Term myTerm = termSet.GetTerms(this.SubstringBefore(sourceValue, "|"), false).FirstOrDefault();

            if (myTerm != null)
            {
                string termString =
                    string.Concat(myTerm.GetDefaultLabel(lcid), TaxonomyField.TaxonomyGuidLabelDelimiter, myTerm.Id);
                int[] ids =
                    TaxonomyField.GetWssIdsOfTerm(targetSPListItem.Web.Site, termStore.Id, termSet.Id, myTerm.Id, true, 1);

                // set the WssId (TaxonomyHiddenList ID) to -1 so that it is added to the TaxonomyHiddenList
                if (ids.Length == 0) termString = "-1;#" + termString;
                else termString = ids[0] + ";#" + termString;

                targetSPListItem[targetSPField.Id] = termString;
            }
        }

        private void MatchSPFieldDefaultId(SPListItem targetSPListItem, SPField targetSPField,
            SPListItem sourceSPListItem, SPField sourceSPField)
        {
            // can't find an obvious transformation path; copy the value as is
            targetSPListItem[targetSPField.Id] = sourceSPListItem[sourceSPField.Id];
        }
        #endregion

        protected string stERR { get; set; }

        private SPFieldLookupValue GetLookupValue(Guid guid, SPFieldLookup lookupField, string p)
        {
            return null;
        }

        private string SubstringAfter(string sourceValue, string p)
        {
            return null;
        }

        private string SubstringBefore(string sourceValue, string p)
        {
            return null;
        }

        private bool FieldIsCopyTarget(SPField targetField, SPListItem sourceItem, ref SPField sourceField)
        {
            // do not copy internal, read-only fields, or the file name
            if (!targetField.ReadOnlyField
                && (targetField.Type != SPFieldType.Attachments)
                && (targetField.Type != SPFieldType.File)
                && (targetField.Type != SPFieldType.Computed)

                && (targetField.Title.Equals("Name") == false)

                && (targetField.Id != SPBuiltInFieldId.DocIcon) //"Type"
                && (targetField.Id != SPBuiltInFieldId.ContentType) //"ContentType"
                && (targetField.Id != SPBuiltInFieldId.ContentTypeId) //"ContentTypeId"
                && (targetField.Id != SPBuiltInFieldId.TemplateUrl) //"Template Link"
                && (targetField.Id != SPBuiltInFieldId.xd_ProgID) //"Html File Link"
                && (targetField.Id != SPBuiltInFieldId.xd_Signature) //"Is Signed"
                && (targetField.Id != SPBuiltInFieldId.MetaInfo) //"Property Bag"
                && (targetField.InternalName != "TaxCatchAll")
                && (targetField.InternalName != "TaxCatchAllLabel")
                && (targetField.Title.Equals("Signatures Status") == false))
            {
                sourceField = sourceItem.Fields.Cast<SPField>().FirstOrDefault(f => f.Title == targetField.Title);

                // source list-item is missing the metadata field
                if (sourceField == null) return false;
                else return true;
            }
            else return false;
        }

        private SPFile GetFileObject(SPSite iSite, string fromFileUrl)
        {
            return null;
        }
    }
}