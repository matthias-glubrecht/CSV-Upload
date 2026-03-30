define([], function() {
  return {
    // Property Pane
    "PropertyPaneDescription": "CSV Upload - Configuration",
    "BasicGroupName": "Settings",
    "DescriptionFieldLabel": "Description",

    // Picker labels
    "SiteCollectionLabel": "Site Collection",
    "WebLabel": "Website",
    "ListLabel": "List",
    "Loading": "Loading...",
    "PleaseSelect": "Please select...",

    // Main headings
    "WebPartTitle": "CSV Data Import",
    "WebPartSubTitle": "Import CSV files into SharePoint lists",
    "SectionTargetTitle": "1. Select target",
    "SectionUploadTitle": "2. Upload CSV file",
    "SectionMappingTitle": "3. Field mapping",

    // Drop zone
    "DropZoneHint": "Drag a CSV file here or click to select",
    "DropZoneHintReplace": "Click or drag another file here",

    // Mapping table headers
    "HeaderKeyColumn": "Key",
    "HeaderSharePointField": "SharePoint Field",
    "HeaderFieldType": "Type",
    "HeaderRequired": "Required",
    "HeaderDefaultValue": "Default Value",
    "HeaderCsvColumn": "CSV Column",
    "HeaderOptions": "Options",

    // Mapping table values
    "Yes": "Yes",
    "No": "No",
    "NotMapped": "-- Not mapped --",
    "DefaultValuePlaceholder": "Default value...",
    "AllowFillInLabel": "Allow fill-in",
    "KeyColumnAriaLabel": "Key column",

    // Field type labels
    "FieldTypeText": "Text",
    "FieldTypeNote": "Multi-line text",
    "FieldTypeNumber": "Number",
    "FieldTypeCurrency": "Currency",
    "FieldTypeDateTime": "Date/Time",
    "FieldTypeChoice": "Choice",
    "FieldTypeMultiChoice": "Choice (multiple)",
    "FieldTypeLookup": "Lookup",
    "FieldTypeLookupMulti": "Lookup (multiple)",
    "FieldTypeUser": "Person",
    "FieldTypeUserMulti": "Person (multiple)",
    "FieldTypeBoolean": "Yes/No",
    "FieldTypeURL": "Hyperlink",
    "FieldTypeTaxonomy": "Managed Metadata",
    "FieldTypeTaxonomyMulti": "Managed Metadata (multiple)",

    // Buttons
    "ImportButtonText": "Import data",
    "ResetButtonText": "Reset",

    // Validation messages
    "ValidationRequiredFields": "\u26A0 Not all required fields are mapped. Please map all required fields to a CSV column or provide a default value.",
    "ValidationNoKeyColumn": "\u2139 No key column selected. Without a key column, all rows will be created as new entries (no updating existing entries).",

    // Progress
    "ImportRunningLabel": "Importing data...",
    "ImportCompletedLabel": "Import completed",
    "ProgressDescription": "{0} of {1} rows processed \u2014 {2} created, {3} updated, {4} errors",
    "ImportSuccessMessage": "All {0} rows were imported successfully. ({1} newly created, {2} updated)",
    "ImportWarningMessage": "Import completed with {0} error(s). ({1} created, {2} updated)",
    "ImportErrorMessage": "The import was aborted due to an error.",
    "ErrorMessagesTitle": "Error messages:",

    // Error messages
    "ErrorLoadingFields": "Error loading list fields: ",
    "ErrorNoHeaders": "The CSV file does not contain column headers.",
    "ErrorReadingFile": "Error reading the file.",
    "ImportAborted": "Import aborted: ",
    "ErrorRowPrefix": "Row {0}: ",
    "ErrorTruncated": "... additional errors are not displayed",

    // Spinner
    "LoadingFieldsLabel": "Loading list fields...",

    // Service error messages
    "ErrorCreatingItem": "Error creating item: ",
    "ErrorUpdatingItem": "Error updating item (ID {0}): ",
    "ErrorSettingTaxonomyField": "Error setting taxonomy field: ",
    "ErrorSettingTaxonomyMultiField": "Error setting multi-value taxonomy field: ",
    "ErrorTaxonomyTermNotFound": "The value \"{0}\" could not be found in the term set for field \"{1}\".",
    "ErrorTaxonomyTermNotInTermSet": "The value \"{0}\" could not be written to field \"{1}\" because it does not occur in the associated term set.",
    "ErrorChoiceValueInvalid": "The value \"{0}\" is not a valid choice for field \"{1}\". Enable \"Allow fill-in\" to allow custom values.",
    "ErrorMultiChoiceValuesInvalid": "The following values are not valid choices for field \"{1}\": {0}. Enable \"Allow fill-in\" to allow custom values.",
    "ErrorLookupValueNotFound": "The value \"{0}\" could not be found in the lookup list for field \"{1}\".",
    "ErrorLookupValuesNotFound": "The following values could not be found in the lookup list for field \"{1}\": {0}.",
    "ErrorUserNotFound": "The user \"{0}\" could not be resolved for field \"{1}\".",
    "ErrorUsersNotFound": "The following users could not be resolved for field \"{1}\": {0}.",

    "FieldErrorDialogTitle": "Value could not be processed",
    "FieldErrorDialogRowLabel": "Row",
    "FieldErrorDialogFieldLabel": "Field",
    "FieldErrorDialogValueLabel": "Value",
    "FieldErrorDialogErrorLabel": "Problem",
    "FieldErrorDialogCorrectedValueLabel": "Corrected value",
    "FieldErrorDialogCorrectedValuePlaceholder": "Enter a new value or leave empty to skip this field...",
    "FieldErrorDialogUseValueButton": "Use corrected value",
    "FieldErrorDialogSkipFieldButton": "Skip field",
    "FieldErrorDialogSkipRowButton": "Skip entire row"
  }
});
