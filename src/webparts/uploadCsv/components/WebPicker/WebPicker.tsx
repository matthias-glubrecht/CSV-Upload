import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IWeb } from '../../models';
import CsvUploadService from '../../service/CsvUploadService';
import * as strings from 'UploadCsvWebPartStrings';

export interface IWebPickerProps {
  service: CsvUploadService;
  siteCollectionUrl: string;
  selectedWeb: IWeb | undefined;
  onWebChanged: (web: IWeb) => void;
}

export interface IWebPickerState {
  options: IDropdownOption[];
  selectedKey: string | undefined;
  loading: boolean;
}

export default class WebPicker extends React.Component<IWebPickerProps, IWebPickerState> {

  constructor(props: IWebPickerProps) {
    super(props);
    this.state = {
      options: [],
      selectedKey: props.selectedWeb ? props.selectedWeb.url : undefined,
      loading: true
    };
  }

  public componentDidMount(): void {
    this._loadWebs(this.props.siteCollectionUrl);
  }

  public componentWillReceiveProps(nextProps: IWebPickerProps): void {
    if (nextProps.siteCollectionUrl !== this.props.siteCollectionUrl) {
      this.setState({ loading: true, options: [], selectedKey: undefined });
      this._loadWebs(nextProps.siteCollectionUrl);
    }
  }

  public render(): React.ReactElement<IWebPickerProps> {
    return (
      <Dropdown
        label={strings.WebLabel}
        placeHolder={this.state.loading ? strings.Loading : strings.PleaseSelect}
        options={this.state.options}
        selectedKey={this.state.selectedKey}
        onChanged={this._onChanged}
        disabled={this.state.loading}
      />
    );
  }

  private _onChanged = (option: IDropdownOption): void => {
    this.setState({ selectedKey: option.key as string });
    // tslint:disable-next-line:no-any
    const optionData: { id: string } = (option as any).data;
    this.props.onWebChanged({
      title: option.text,
      url: option.key as string,
      id: optionData ? optionData.id : ''
    });
  }

  private _loadWebs(siteCollectionUrl: string): void {
    this.props.service.getWebs(siteCollectionUrl).then((webs: IWeb[]) => {
      const options: IDropdownOption[] = webs.map((web: IWeb) => ({
        key: web.url,
        text: `${web.title} (${web.url})`,
        data: { id: web.id }
      }));

      // Auto-select: if current web belongs to this site collection, use it; otherwise root web
      const currentWeb: IWeb = this.props.service.getCurrentWeb();
      let selectedKey: string | undefined = undefined;

      if (this.props.selectedWeb) {
        selectedKey = this.props.selectedWeb.url;
      } else if (currentWeb.url.toLowerCase().indexOf(siteCollectionUrl.toLowerCase()) === 0) {
        selectedKey = currentWeb.url;
      } else if (options.length > 0) {
        selectedKey = options[0].key as string;
      }

      this.setState({ options: options, selectedKey: selectedKey, loading: false });

      // Notify parent of auto-selection
      if (selectedKey) {
        const selectedOption: IDropdownOption | undefined = options.filter(
          (o: IDropdownOption) => o.key === selectedKey
        )[0];
        if (selectedOption) {
          const selectedWeb: IWeb = webs.filter((w: IWeb) => w.url === selectedKey)[0];
          if (selectedWeb) {
            this.props.onWebChanged(selectedWeb);
          }
        }
      }
    });
  }
}
