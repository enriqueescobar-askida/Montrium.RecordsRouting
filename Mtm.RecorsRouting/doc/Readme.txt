RecordCentreManager - Manages all the resources from a record centre.
Needs:
Has:
-the Url to connect to the record centre.
-the list of files in the record centre.
-the SPListCollection of lists from SPWeb.
-the SPLsitItemCollection of RecordsList from the Drop Off Library.
Does:
-Connects to the SPSite using the Url.
-Connects to the SPWeb using the SPSite.
-Overrides ToString()

RecordsManager - Manages all the resources from a set of records in the record centre.
Needs:
Has:
-a List<RecordCentreDocument>
Does:
-Fetches in SPListItemCollection the files to create a list of RecordCentreDocuments with RCListCollection.
-Overrides ToString().

RecordCentreDocuments - Manages all the resources from a record centre document.
Needs:
Has:
-a bool IsOrphan
-a bool IsAdoptable
-an int Matches
-a string Path
-a string ContentTypeLibrary
-a string Author
-a string ModifiedBy
-a string Title
-a string ContentTypeId
-a string XmlProperties
-a SPList SPListCandidate
-a SPListItem SpListItem
-a SPFieldCollection SpFieldCollection
-a SPContentType SPContentType
Does:
-Fetches all info inside the internal RecordCentre.
-Overrides ToString().

SPFieldManager - Manages all the resources
Needs:
-SPListItem fileSpListItem
-SPList contextualSPList
-string url
-string xmlLookup
-bool changeVersion
Has:
-a string ContextUrl
-a string ContextListName
-a SPListItem FileSPListItem
-a SPFieldCollection ContextSPFields
-a SPSite ContextSPSite
-a SPWeb ContextSPWeb
-a bool ChangeVersion
Does:
-Overrides ToString().

SPFieldUpdater - Manages all the resources to update and SPField.
Needs:
-SPField contextSPField
-SPField fileSPField
Has:
-a bool SPFieldUserIsMultiValue
-a bool SPFieldLookupIsMultiValue
-a bool SPFieldTypeIsTextOrChoice
-a bool SPFieldTypeIsNumber
-a bool SPFieldTypeHasSameUser
-a bool SPFieldTypeIsUser
-a bool SPFieldTypeIsLookup
-a bool SPFieldTypeIsLookupOrInvalid
-a bool SPFieldTypeIsDateTime
-a bool SPFieldIsTaxonomy
-a SPFieldLookup FileSPFieldLookup
-a SPFieldUser FileSPFieldUser
-a SPField FileSPField
-a SPField ContextSPField
Does:
-Fixes the fields following the screeing of the SPField information type.
-Overrides ToString().

XmlLookupReader - Manages all the resources to read an XML Lookup.
Needs:
-string xmlProperties
Has:
-a List<XmlLookupNode> with valid node containing not empty values.
-a XmlLookupNode for the matched information.
Does:
-Parses an XML Lookup and creates a list of XmlLookupNodes.
-Overrides ToString().

XmlLookupNode - Manages all the resources of a node in the XML Lookup.
Needs:
-string fieldName to validate or to trim
-string fieldValue
-string fieldType
Has:
-a name.
-a type.
-a value.
Does:
-Fetches all the info inside a XML property from the properties of the Lookup.
-Overrides ToString().

