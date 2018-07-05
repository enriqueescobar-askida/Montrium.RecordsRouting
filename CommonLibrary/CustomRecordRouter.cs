// -----------------------------------------------------------------------
// <copyright file="CustomRecordRouter.cs" company="Montrium">
// IT LIcense
// </copyright>
// -----------------------------------------------------------------------

namespace Mtm.RecordsRouting.CommonLibrary
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;

    using Microsoft.Office.RecordsManagement.RecordsRepository;
    using Microsoft.SharePoint;

    using RecordsRepositoryProperty = Microsoft.SharePoint.RecordsRepositoryProperty;

    /// <summary>Custom RecordRouter.</summary>
    public class CustomRecordRouter : ICustomRouter
    {
        #region Constructor
        #endregion

        #region Destructor
        #endregion

        #region PublicMethods
        /// <summary>On SubmitFile.</summary>
        /// <param name="contentOrganizerWeb">
        /// The content organizer web. Web site to which the document is being added.
        /// </param>
        /// <param name="recordSeries">
        /// The record series. Content type of the document.
        /// </param>
        /// <param name="userName">
        /// The user name. Login name of the user creating the file.
        /// </param>
        /// <param name="fileContent">
        /// The file content. Content stream of the file being organized.
        /// </param>
        /// <param name="properties">
        /// The properties. Metadata of the file being organized.
        /// </param>
        /// <param name="finalFolder">
        /// The final folder. Final location configured for the document being organized.
        /// </param>
        /// <param name="resultDetails">
        /// The result details. Any details that the custom router wants to furnish for logging purposes.
        /// </param>
        /// <returns>Custom RouterResult. Custom information that should be logged by the content organizer.</returns>
        public CustomRouterResult OnSubmitFile(
            EcmDocumentRoutingWeb contentOrganizerWeb,
            string recordSeries,
            string userName,
            Stream fileContent,
            RecordsRepositoryProperty[] properties,
            SPFolder finalFolder,
            ref string resultDetails)
        {
            if (contentOrganizerWeb == null) throw new ArgumentNullException("contentOrganizerWeb");
            // We should have a Content Organizer enabled web 
            if (!contentOrganizerWeb.IsRoutingEnabled) throw new ArgumentException("Invalid Content Organizer that invoked the custom router.");
            if (String.IsNullOrEmpty(recordSeries)) throw new ArgumentNullException("Invalid Content type of the document.");
            if (String.IsNullOrEmpty(userName)) throw new ArgumentNullException("Invalid Login name of the user creating the file.");
            if (fileContent == null) throw new IOException("Invalid Content stream of the file being organized.");
            if (properties == null) throw new SPFieldValidationException("Invalid Metadata of the file being organized.");
            if (finalFolder == null) throw new DirectoryNotFoundException("Invalid Final location configured for the document being organized.");
            if (String.IsNullOrEmpty(resultDetails)) throw new ArgumentNullException("Invalid Custom information that should be logged by the content organizer.");

            return CustomRouterResult.SuccessCancelFurtherProcessing;
        }
        #endregion

        #region PrivateMethods
        /// <summary>Find Folder. </summary>
        /// <param name="name">The name.</param>
        /// <param name="sPFolderCollection">The SP Folder Collection container.</param>
        /// <returns>A SPFolder</returns>
        private SPFolder FindFolder(string name, SPFolderCollection sPFolderCollection)
        {
            foreach (SPFolder spFolder in sPFolderCollection)
                if (spFolder.Name == name)
                    return spFolder;

            return null;
        }

        /// <summary>Ok ToCopyField.</summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="spListItem">
        /// The sp list item.
        /// </param>
        /// <returns>A Boolean</returns>
        private bool OkToCopyField(string name, SPListItem spListItem)
        {
            try
            {
                SPField spField = spListItem.Fields.GetField(name);

                // Field does not exist in the destination content type
                // Can't copy readonly fields
                // Can't chenge the content type
                // Other fields types which cannot be copied
                if (spField == null || spField.ReadOnlyField || spField.InternalName == "ContentType" ||
                    spField.Type == SPFieldType.Invalid || spField.Type == SPFieldType.WorkflowStatus ||
                    spField.Type == SPFieldType.File || spField.Type == SPFieldType.Computed) return false;
            }
            catch
            {
                return false;
            }

            return true;
        }

        /// <summary>GetInternal FieldNames.</summary>
        /// <param name="fileUrl">
        /// The file url.
        /// </param>
        /// <param name="fieldNames">
        /// The field names.
        /// </param>
        /// <returns>Dictionary dict.</returns>
        private Dictionary<string, string> GetInternalFieldNames(string fileUrl,
            List<string> fieldNames)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            try
            {
                // try and obtain a valid instance of the SPFile class at the specified URL
                object iSpObject = SPContext.Current.Site.RootWeb.GetFileOrFolderObject(fileUrl);
                SPFile spFile = null;

                if (iSpObject is SPFile)
                    spFile = iSpObject as SPFile;

                if (spFile != null)
                {
                    SPListItem spListItem = spFile.Item;

                    foreach (string fieldName in fieldNames)
                    {
                        try
                        {
                            SPField spField = spListItem.Fields.GetField(fieldName);
                            if (spField != null)
                                dict.Add(fieldName, spField.InternalName);
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

        /// <summary>SP Folder.</summary>
        /// <param name="sStudyNumber">The s study number.</param>
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
        private SPFolder GetDestinationFolder(
            string sStudyNumber, string sSection,
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
