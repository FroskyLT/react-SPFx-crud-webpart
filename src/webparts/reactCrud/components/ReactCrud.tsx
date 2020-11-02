import * as React from 'react';
import styles from './ReactCrud.module.scss';
import { IReactCrudProps } from './IReactCrudProps';
import { IReactCrudState } from './IReactCrudState';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './IListItem';

export default class ReactCrud extends React.Component<IReactCrudProps, IReactCrudState> {
  constructor(props: IReactCrudProps, state: IReactCrudState) {
    super(props);
    this.state = {
      status: "Ready",
      getHeaders: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
    };
  }

  public render(): React.ReactElement<IReactCrudProps> {
    return (
      <div className={styles.reactCrud}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>SharePoint CRUD operations webpart</span>
              <p className={styles.subTitle}>Your Sharepoint list title: {escape(this.props.listTitle)}</p>
              <p className={styles.description}>Status is: {this.state.status}</p>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.customBody}>
              <div className={styles.customRow}>
                <div className={styles.customColumn}>
                  <PrimaryButton text={"Create"} onClick={() => this.createItem()} />
                  <PrimaryButton text={"Read"} onClick={() => this.readItem()} />
                </div>
              </div>
              <div className={styles.customRow}>
                <div className={styles.customColumn}>
                  <PrimaryButton text={"Update"} onClick={async () => this.updateItem()} />
                  <PrimaryButton text={"Delete"} onClick={() => this.deleteItem()} />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private async getLatestItemId(): Promise<number> {

    const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items?$orderby=Id desc&$top=1&$select=id`;

    const response: SPHttpClientResponse = await this.props.spHttpClient
      .get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
    const data: { value: { Id: number }[] } = await response.json();

    if (data.value.length === 0) return -1;
    else return data.value[0].Id;
  }
  private async getLatestItem(): Promise<IListItem> {
    try {
      const itemID: number = await this.getLatestItemId();

      if (itemID === -1) throw new Error('No items found in list');
      this.setState({ status: `Loading information about item with id: ${itemID}...` });

      const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemID})?$select=Title,Id`;

      const response: SPHttpClientResponse = await this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
      const item: IListItem = await response.json();

      return item;

    } catch (error) { this.setState({ status: `Error to read an item: ${error}` }); }
  }


  private createItem(): void {

    this.setState({ status: 'Creating an item...' });

    const body: string = JSON.stringify({ 'Title': 'new-item' });
    const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items`;
    const createHeaders: HeadersInit =
      { 'Content-type': 'application/json;odata=nometadata', ...this.state.getHeaders };

    const postItem: () => Promise<IListItem> = async () => {
      const response: SPHttpClientResponse = await this.props.spHttpClient
        .post(listUrl, SPHttpClient.configurations.v1, { headers: createHeaders, body: body });
      const item: IListItem = await response.json();

      return item;
    };

    postItem().then((item: IListItem): void =>
      this.setState({ status: `Item "${item.Title}" "${item.Id}" successfully created` }),
      (error: any): void => this.setState({ status: `Error to create an item: ${error.message}` })
    );
  }

  private readItem(): void {

    this.setState({ status: 'Loading latest item...' });

    this.getLatestItem().then((item: IListItem) =>
      this.setState({ status: `Item Id: ${item.Id}, Title: ${item.Title}` }));
  }

  private updateItem(): void {

    this.setState({ status: 'Loading latest item...' });

    let latestItemId: number;
    const updateHeaders: HeadersInit =
      { ...this.state.getHeaders, 'Content-type': 'application/json;odata=nometadata', 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' };

     const update: () => void = async () => {
      try {
        const item: IListItem = await this.getLatestItem();

        if(typeof item === 'undefined') throw new Error('while reading an item');
          latestItemId = item.Id;
          this.setState({ status: 'Pending to update item...' });
          const body: string = JSON.stringify({ 'Title': 'updated-item' });
          const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;
          
          const response: SPHttpClientResponse = await this.props.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, { headers: updateHeaders, body: body });
          
          this.setState({ status: `Item with Id: ${latestItemId} successfully updated` });


      } catch (error) { this.setState({ status: `Error to update an item: ${error.message}` }); }
     };

     update();
  }

  private deleteItem(): void {
    this.setState({ status: 'Loading latest item...' });

    let latestItemId: number;
    let etag: string;

    const getLItem: () => Promise<IListItem> = async () => {
      try {
        const itemID: number = await this.getLatestItemId();
        
        if (itemID === -1) throw new Error('No items found in list');
        this.setState({ status: `Loading information about item with id: ${itemID}...` });
        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemID})?$select=Title,Id`;

        const response: SPHttpClientResponse = await this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
        const item: IListItem = await response.json();
        etag = response.headers.get('ETag');
        latestItemId = item.Id;
        return item;

      } catch (error) { this.setState({ status: `Error to read an item: ${error}` }); }
    };

    const deleteLItem: () => void = async () => {
      try {
        const item: IListItem = await getLItem();
        if(typeof item === 'undefined') throw new Error('while reading an item');

        this.setState({ status: `Deleting item with Id: ${item.Id}...` });
        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;
        const deleteHeaders: HeadersInit =
          { ...this.state.getHeaders, 'Content-type': 'application/json;odata=nometadata', 'IF-MATCH': etag, 'X-HTTP-Method': 'DELETE' };

        const response: SPHttpClientResponse = await this.props.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, { headers: deleteHeaders });

        this.setState({ status: `Item with Id: ${latestItemId} successfully deleted` });

      } catch (error) { this.setState({ status: `Error to delete an item: ${error.message}` }); }
    };
    deleteLItem();
  }
}
