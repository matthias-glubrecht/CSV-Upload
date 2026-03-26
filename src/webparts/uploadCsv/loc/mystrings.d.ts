declare interface IUploadCsvWebPartStrings {
  // Property Pane
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  // Picker labels
  SiteCollectionLabel: string;
  WebLabel: string;
  ListLabel: string;
  Loading: string;
  PleaseSelect: string;

  // Main headings
  WebPartTitle: string;
  WebPartSubTitle: string;
  SectionTargetTitle: string;
  SectionUploadTitle: string;
  SectionMappingTitle: string;

  // Drop zone
  DropZoneHint: string;
  DropZoneHintReplace: string;

  // Mapping table headers
  HeaderKeyColumn: string;
  HeaderSharePointField: string;
  HeaderFieldType: string;
  HeaderRequired: string;
  HeaderDefaultValue: string;
  HeaderCsvColumn: string;
  HeaderOptions: string;

  // Mapping table values
  Yes: string;
  No: string;
  NotMapped: string;
  DefaultValuePlaceholder: string;
  AllowFillInLabel: string;
  KeyColumnAriaLabel: string;

  // Field type labels
  FieldTypeText: string;
  FieldTypeNote: string;
  FieldTypeNumber: string;
  FieldTypeCurrency: string;
  FieldTypeDateTime: string;
  FieldTypeChoice: string;
  FieldTypeMultiChoice: string;
  FieldTypeLookup: string;
  FieldTypeLookupMulti: string;
  FieldTypeUser: string;
  FieldTypeUserMulti: string;
  FieldTypeBoolean: string;
  FieldTypeURL: string;
  FieldTypeTaxonomy: string;
  FieldTypeTaxonomyMulti: string;

  // Buttons
  ImportButtonText: string;
  ResetButtonText: string;

  // Validation messages
  ValidationRequiredFields: string;
  ValidationNoKeyColumn: string;

  // Progress
  ImportRunningLabel: string;
  ImportCompletedLabel: string;
  ProgressDescription: string;
  ImportSuccessMessage: string;
  ImportWarningMessage: string;
  ImportErrorMessage: string;
  ErrorMessagesTitle: string;

  // Error messages
  ErrorLoadingFields: string;
  ErrorNoHeaders: string;
  ErrorReadingFile: string;
  ImportAborted: string;
  ErrorRowPrefix: string;
  ErrorTruncated: string;

  // Spinner
  LoadingFieldsLabel: string;

  // Service error messages
  ErrorCreatingItem: string;
  ErrorUpdatingItem: string;
  ErrorSettingTaxonomyField: string;
  ErrorSettingTaxonomyMultiField: string;
}

declare module 'UploadCsvWebPartStrings' {
  const strings: IUploadCsvWebPartStrings;
  export = strings;
}
