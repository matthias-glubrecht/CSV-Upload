define([], function() {
  return {
    // Property Pane
    "PropertyPaneDescription": "CSV Upload - Konfiguration",
    "BasicGroupName": "Einstellungen",
    "DescriptionFieldLabel": "Beschreibung",

    // Picker labels
    "SiteCollectionLabel": "Websitesammlung",
    "WebLabel": "Website",
    "ListLabel": "Liste",
    "Loading": "Laden...",
    "PleaseSelect": "Bitte auswählen...",

    // Main headings
    "WebPartTitle": "CSV-Datenimport",
    "WebPartSubTitle": "CSV-Dateien in SharePoint-Listen importieren",
    "SectionTargetTitle": "1. Ziel auswählen",
    "SectionUploadTitle": "2. CSV-Datei hochladen",
    "SectionMappingTitle": "3. Feldzuordnung",

    // Drop zone
    "DropZoneHint": "CSV-Datei hierher ziehen oder klicken zum Auswählen",
    "DropZoneHintReplace": "Klicken oder eine andere Datei hierher ziehen",

    // Mapping table headers
    "HeaderKeyColumn": "Schlüssel",
    "HeaderSharePointField": "SharePoint-Feld",
    "HeaderFieldType": "Typ",
    "HeaderRequired": "Pflichtfeld",
    "HeaderDefaultValue": "Defaultwert",
    "HeaderCsvColumn": "CSV-Spalte",
    "HeaderOptions": "Optionen",

    // Mapping table values
    "Yes": "Ja",
    "No": "Nein",
    "NotMapped": "-- Nicht zugeordnet --",
    "DefaultValuePlaceholder": "Defaultwert...",
    "AllowFillInLabel": "Freie Eingabe",
    "KeyColumnAriaLabel": "Schlüsselspalte",

    // Field type labels
    "FieldTypeText": "Text",
    "FieldTypeNote": "Mehrzeiliger Text",
    "FieldTypeNumber": "Zahl",
    "FieldTypeCurrency": "Währung",
    "FieldTypeDateTime": "Datum/Uhrzeit",
    "FieldTypeChoice": "Auswahl",
    "FieldTypeMultiChoice": "Auswahl (mehrere)",
    "FieldTypeLookup": "Nachschlagen",
    "FieldTypeLookupMulti": "Nachschlagen (mehrere)",
    "FieldTypeUser": "Person",
    "FieldTypeUserMulti": "Person (mehrere)",
    "FieldTypeBoolean": "Ja/Nein",
    "FieldTypeURL": "Hyperlink",
    "FieldTypeTaxonomy": "Verwaltete Metadaten",
    "FieldTypeTaxonomyMulti": "Verwaltete Metadaten (mehrere)",

    // Buttons
    "ImportButtonText": "Daten importieren",
    "ResetButtonText": "Zurücksetzen",

    // Validation messages
    "ValidationRequiredFields": "⚠ Es sind nicht alle Pflichtfelder zugeordnet. Bitte ordnen Sie alle Pflichtfelder einer CSV-Spalte zu oder geben Sie einen Defaultwert an.",
    "ValidationNoKeyColumn": "ℹ Kein Schlüsselfeld ausgewählt. Ohne Schlüsselfeld werden alle Zeilen als neue Einträge erstellt (kein Update bestehender Einträge).",

    // Progress
    "ImportRunningLabel": "Daten werden importiert...",
    "ImportCompletedLabel": "Import abgeschlossen",
    "ProgressDescription": "{0} von {1} Zeilen verarbeitet — {2} erstellt, {3} aktualisiert, {4} Fehler",
    "ImportSuccessMessage": "Alle {0} Zeilen wurden erfolgreich importiert. ({1} neu erstellt, {2} aktualisiert)",
    "ImportWarningMessage": "Import abgeschlossen mit {0} Fehler(n). ({1} erstellt, {2} aktualisiert)",
    "ImportErrorMessage": "Der Import wurde aufgrund eines Fehlers abgebrochen.",
    "ErrorMessagesTitle": "Fehlermeldungen:",

    // Error messages
    "ErrorLoadingFields": "Fehler beim Laden der Listenfelder: ",
    "ErrorNoHeaders": "Die CSV-Datei enthält keine Spaltenüberschriften.",
    "ErrorReadingFile": "Fehler beim Lesen der Datei.",
    "ImportAborted": "Import abgebrochen: ",
    "ErrorRowPrefix": "Zeile {0}: ",
    "ErrorTruncated": "... weitere Fehler werden nicht angezeigt",

    // Spinner
    "LoadingFieldsLabel": "Listenfelder werden geladen...",

    // Service error messages
    "ErrorCreatingItem": "Fehler beim Erstellen: ",
    "ErrorUpdatingItem": "Fehler beim Aktualisieren (ID {0}): ",
    "ErrorSettingTaxonomyField": "Fehler beim Setzen des Taxonomie-Feldes: ",
    "ErrorSettingTaxonomyMultiField": "Fehler beim Setzen des Taxonomie-Multi-Feldes: "
  }
});
