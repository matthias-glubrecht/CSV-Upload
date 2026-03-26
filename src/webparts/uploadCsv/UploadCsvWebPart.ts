import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'UploadCsvWebPartStrings';
import UploadCsv from './components/UploadCsv';
import { IUploadCsvProps } from './components/IUploadCsvProps';
import CsvUploadService from './service/CsvUploadService';

export interface IUploadCsvWebPartProps {
  description: string;
}

export default class UploadCsvWebPart extends BaseClientSideWebPart<IUploadCsvWebPartProps> {
  private _service: CsvUploadService;

  public render(): void {
    const element: React.ReactElement<IUploadCsvProps> = React.createElement(
      UploadCsv,
      {
        description: this.properties.description,
        service: this._service
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._service = new CsvUploadService(this.context);
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
