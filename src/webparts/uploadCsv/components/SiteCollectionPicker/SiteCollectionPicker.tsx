import * as React from 'react';
import { ComboBox, IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { ISiteCollection } from '../../models';
import CsvUploadService from '../../service/CsvUploadService';
import * as strings from 'UploadCsvWebPartStrings';

export interface ISiteCollectionPickerProps {
  service: CsvUploadService;
  selectedSiteCollection: ISiteCollection | undefined;
  onSiteCollectionChanged: (siteCollection: ISiteCollection) => void;
}

export interface ISiteCollectionPickerState {
  options: IComboBoxOption[];
  selectedKey: string | undefined;
  searchText: string;
}

// tslint:disable-next-line:max-line-length
export default class SiteCollectionPicker extends React.Component<ISiteCollectionPickerProps, ISiteCollectionPickerState> {
  private _searchTimeout: number;

  constructor(props: ISiteCollectionPickerProps) {
    super(props);
    const currentSite: ISiteCollection = props.service.getCurrentSiteCollection();
    this.state = {
      options: [{
        key: currentSite.url,
        text: `${currentSite.title} (${currentSite.url})`
      }],
      selectedKey: props.selectedSiteCollection ? props.selectedSiteCollection.url : currentSite.url,
      searchText: ''
    };
  }

  public render(): React.ReactElement<ISiteCollectionPickerProps> {
    return (
      <ComboBox
        label={strings.SiteCollectionLabel}
        autoComplete='on'
        allowFreeform={true}
        selectedKey={this.state.selectedKey}
        options={this.state.options}
        onChanged={this._onChanged}
      />
    );
  }

  private _onChanged = (option: IComboBoxOption, index?: number, value?: string): void => {
    if (option) {
      this.setState({ selectedKey: option.key as string });
      this.props.onSiteCollectionChanged({
        title: option.text,
        url: option.key as string
      });
    } else if (value && value.length >= 2) {
      // User typed freeform text - search for matching site collections
      if (this._searchTimeout) {
        window.clearTimeout(this._searchTimeout);
      }
      this._searchTimeout = window.setTimeout(() => {
        this.props.service.searchSiteCollections(value).then((sites: ISiteCollection[]) => {
          const options: IComboBoxOption[] = sites.map((site: ISiteCollection) => ({
            key: site.url,
            text: `${site.title} (${site.url})`
          }));
          this.setState({ options: options });
        });
      }, 300);
    }
  }
}
