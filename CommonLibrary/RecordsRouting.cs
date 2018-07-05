// -----------------------------------------------------------------------
// <copyright file="RecordsRouting.cs" company="Montrium">
// MIT Licence
// </copyright>
// -----------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;

    using Microsoft.Office.RecordsManagement.RecordsRepository;
    using Microsoft.SharePoint;

    using RecordsRepositoryProperty = Microsoft.Office.RecordsManagement.RecordsRepository.RecordsRepositoryProperty;

    /// <summary>Records Routing.</summary>
    public class RecordsRouting : IRouter
    {
        /// <summary>RR RouterResult.</summary>
        /// <param name="recordSeries">The record series.</param>
        /// <param name="sourceUrl">The source url.</param>
        /// <param name="userName">The user name.</param>
        /// <param name="fileToSubmit">The file to submit.</param>
        /// <param name="properties">The properties.</param>
        /// <param name="destination">The destination.</param>
        /// <param name="resultDetails">The result details.</param>
        public RouterResult OnSubmitFile(
            string recordSeries,
            string sourceUrl,
            string userName,
            ref byte[] fileToSubmit,
            ref RecordsRepositoryProperty[] properties,
            ref SPList destination,
            ref string resultDetails)
        {
            try
            {
                string sContentType = string.Empty;

                string sStudyNumber = string.Empty;
                string sMolecule = string.Empty;
                string sSection = string.Empty;
                string sSubSection = string.Empty;
                string sDocumentType = string.Empty;

                string sDocumentSource = string.Empty;

                List<string> recordFields = new List<string>
                    {
                        "Study Number", "Molecule", "Section", "SubSection", "Document Type Name" 
                    };

                // add the properties to a dictionary for easier access.
                Dictionary<string, string> recordProperties = new Dictionary<string, string>();
                foreach (RecordsRepositoryProperty property in properties)
                {
                    if (!recordProperties.ContainsKey(property.Name))
                        recordProperties.Add(property.Name, property.Value);
                }

                if (!recordProperties.ContainsKey("ContentType"))
                {
                    // log the fact we didn't find it...but let the default processing continue
                    Trace.WriteLine("RecordsRouting Failed... ContentType not found");
                    return RouterResult.SuccessContinueProcessing;
                }
                else
                {
                    sContentType = recordProperties["ContentType"].ToString(CultureInfo.InvariantCulture);
                    sDocumentSource = recordProperties["FileRef"].ToString(CultureInfo.InvariantCulture);

                    Dictionary<string, string> internalNames = this.GetInternalFieldNames(sDocumentSource, recordFields);

                    if (recordProperties.ContainsKey(internalNames["Study Number"]))
                    {
                        sStudyNumber = recordProperties[internalNames["Study Number"]].ToString(CultureInfo.InvariantCulture);
                        int index = sStudyNumber.IndexOf(";#");

                        if (index > 0)
                            sStudyNumber = sStudyNumber.Substring(index + 2);
                    }

                    if (recordProperties.ContainsKey(internalNames["Molecule"]))
                    {
                        sMolecule = recordProperties[internalNames["Molecule"]].ToString(CultureInfo.InvariantCulture);
                        int index = sMolecule.IndexOf(";#");

                        if (index > 0)
                            sMolecule = sMolecule.Substring(index + 2);
                    }

                    if (recordProperties.ContainsKey(internalNames["Section"]))
                    {
                        sSection = recordProperties[internalNames["Section"]].ToString();
                        int index = sSection.IndexOf(";#");

                        if (index > 0)
                            sSection = sSection.Substring(index + 2);
                    }

                    if (recordProperties.ContainsKey(internalNames["SubSection"]))
                    {
                        sSubSection = recordProperties[internalNames["SubSection"]].ToString();
                        int index = sSubSection.IndexOf(";#");

                        if (index > 0)
                            sSubSection = sSubSection.Substring(index + 2);
                    }

                    if (recordProperties.ContainsKey(internalNames["Document Type Name"]))
                        sDocumentType = recordProperties[internalNames["Document Type Name"]].ToString(CultureInfo.InvariantCulture);
                }

                // create a new filename using the date & time
                string sFileName = Path.GetFileNameWithoutExtension(sourceUrl) + " ("
                                   + DateTime.Now.ToUniversalTime().ToString("yyMMddHHmmss") + ")"
                                   + Path.GetExtension(sourceUrl);

                // connection info and context
                SPWeb recordCenter = destination.ParentWeb;
                SPList recordLibrary = recordCenter.Lists[sMolecule];

                SPFolder oFolder = this.GetDestinationFolder(
                    sStudyNumber, sSection, sSubSection, sDocumentType, recordLibrary);

                // Add the document
                SPFile oFile = oFolder.Files.Add(sFileName, fileToSubmit);
                SPListItem oItem = oFile.Item;

                if (null == recordLibrary.ContentTypes[sContentType])
                    recordLibrary.ContentTypes.Add(recordCenter.AvailableContentTypes[sContentType]);

                oItem["ContentTypeId"] = recordLibrary.ContentTypes[sContentType].Id;
                oItem["Document Source"] = sDocumentSource;
                oItem.Update();

                foreach (RecordsRepositoryProperty p in properties)
                {
                    if (oItem.Fields.ContainsField(p.Name))
                    {
                        try
                        {
                            if (this.OkToCopyField(p.Name, oItem))
                                oItem[p.Name] = p.Value;
                        }
                        catch (Exception ex)
                        {
                            Trace.WriteLine(string.Format("Failed to copy field '{0}': {1}", p.Name, ex.Message));
                        }
                    }
                }

                oItem.Update();

                return RouterResult.SuccessCancelFurtherProcessing;
            }
            catch (Exception ex)
            {
                resultDetails = "Failed to route record: " + ex.Message;
                Trace.WriteLine(resultDetails);

                return RouterResult.RejectFile;
            }
        }

        #region PrivateMethods
        /// <summary>Find Folder.
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="oContainer">
        /// The o container.
        /// </param>
        /// <returns>A SPFolder
        /// </returns>
        private SPFolder FindFolder(string name, SPFolderCollection oContainer)
        {
            foreach (SPFolder f in oContainer)
                if (f.Name == name)
                    return f;

            return null;
        }

        /// <summary>Is OkToCopyField?
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="oItem">
        /// The o item.
        /// </param>
        /// <returns> A Boolean.
        /// </returns>
        private bool OkToCopyField(string name, SPListItem oItem)
        {
            try
            {
                SPField oField = oItem.Fields.GetField(name);

                // Field does not exist in the destination content type
                // Can't copy readonly fields
                // Can't chenge the content type
                // Other fields types which cannot be copied
                if (oField == null || oField.ReadOnlyField || oField.InternalName == "ContentType" ||
                    oField.Type == SPFieldType.Invalid || oField.Type == SPFieldType.WorkflowStatus ||
                    oField.Type == SPFieldType.File || oField.Type == SPFieldType.Computed) return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        /// <summary>GetInternal FieldNames.
        /// </summary>
        /// <param name="fileUrl">
        /// The file url.
        /// </param>
        /// <param name="fieldNames">
        /// The field names.
        /// </param>
        /// <returns>Dictionary dict.
        /// </returns>
        private Dictionary<string, string> GetInternalFieldNames(string fileUrl, List<string> fieldNames)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            try
            {
                // try and obtain a valid instance of the SPFile class at the specified URL
                object iSpObject = SPContext.Current.Site.RootWeb.GetFileOrFolderObject(fileUrl);
                SPFile iFile = iSpObject as SPFile;

                if (iFile != null)
                {
                    SPListItem iItem = iFile.Item;
                    foreach (string fieldName in fieldNames)
                    {
                        try
                        {
                            SPField field = iItem.Fields.GetField(fieldName);
                            if (field != null)
                                dict.Add(fieldName, field.InternalName);
                            else
                                dict.Add(fieldName, string.Empty);
                        }
                        catch
                        {
                            dict.Add(fieldName, string.Empty);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }

            return dict;
        }

        /// <summary>SP Folder.
        /// </summary>
        /// <param name="sStudyNumber">
        /// The s study number.
        /// </param>
        /// <param name="sSection">
        /// The s section.
        /// </param>
        /// <param name="sSubSection">
        /// The s sub section.
        /// </param>
        /// <param name="sDocumentType">
        /// The s document type.
        /// </param>
        /// <param name="destination">
        /// The destination.
        /// </param>
        /// <returns>SPFolder destinationFolder.
        /// </returns>
        private SPFolder GetDestinationFolder(string sStudyNumber, string sSection,
            string sSubSection, string sDocumentType, SPList destination)
        {
            SPFolder studyFolder = this.FindFolder(sStudyNumber, destination.RootFolder.SubFolders);
            if (studyFolder == null)
            {
                if (String.IsNullOrEmpty(sStudyNumber))
                {
                    destination.RootFolder.SubFolders.Add("General Clinical Files");
                    studyFolder = this.FindFolder("General Clinical Files", destination.RootFolder.SubFolders);
                }
                else
                {
                    destination.RootFolder.SubFolders.Add(sStudyNumber);
                    studyFolder = this.FindFolder(sStudyNumber, destination.RootFolder.SubFolders);
                }
            }

            SPFolder sectionFolder = this.FindFolder(sSection, studyFolder.SubFolders);
            if (sectionFolder == null)
            {
                studyFolder.SubFolders.Add(sSection);
                sectionFolder = this.FindFolder(sSection, studyFolder.SubFolders);
            }

            SPFolder subSectionFolder = this.FindFolder(sSubSection, sectionFolder.SubFolders);
            if (subSectionFolder == null)
            {
                sectionFolder.SubFolders.Add(sSubSection);
                subSectionFolder = this.FindFolder(sSubSection, sectionFolder.SubFolders);
            }

            SPFolder destinationFolder = this.FindFolder(sDocumentType, subSectionFolder.SubFolders);
            if (destinationFolder == null)
            {
                subSectionFolder.SubFolders.Add(sDocumentType);
                destinationFolder = this.FindFolder(sDocumentType, subSectionFolder.SubFolders);
            }

            return destinationFolder;
        }
        #endregion
    }
}
