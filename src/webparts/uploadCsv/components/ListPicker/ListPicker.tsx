import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { IListInfo } from '../../models';
import CsvUploadService from '../../service/CsvUploadService';
import * as strings from 'UploadCsvWebPartStrings';

export interface IListPickerProps {
  service: CsvUploadService;
  webUrl: string;
  selectedList: IListInfo | undefined;
  onListChanged: (list: IListInfo) => void;
}

export interface IListPickerState {
  options: IDropdownOption[];
  selectedKey: string | undefined;
  loading: boolean;
}

export default class ListPicker extends React.Component<IListPickerProps, IListPickerState> {

  constructor(props: IListPickerProps) {
    super(props);
    this.state = {
      options: [],
      selectedKey: props.selectedList ? props.selectedList.id : undefined,
      loading: true
    };
  }

  public componentDidMount(): void {
    this._loadLists(this.props.webUrl);
  }

  public componentWillReceiveProps(nextProps: IListPickerProps): void {
    if (nextProps.webUrl !== this.props.webUrl) {
      this.setState({ loading: true, options: [], selectedKey: undefined });
      this._loadLists(nextProps.webUrl);
    }
  }

  public render(): React.ReactElement<IListPickerProps> {
    return (
      <Dropdown
        label={strings.ListLabel}
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
    this.props.onListChanged({
      title: option.text,
      id: option.key as string
    });
  }

  private _loadLists(webUrl: string): void {
    this.props.service.getLists(webUrl).then((lists: IListInfo[]) => {
      const options: IDropdownOption[] = lists.map((list: IListInfo) => ({
        key: list.id,
        text: list.title
      }));
      this.setState({ options: options, loading: false });
    });
  }
}
