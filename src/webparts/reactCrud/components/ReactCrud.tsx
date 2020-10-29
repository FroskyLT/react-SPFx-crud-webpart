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
      status: "Ready"
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
                  <PrimaryButton text={"Update"} onClick={() => this.updateItem()} />
                  <PrimaryButton text={"Delete"} onClick={() => this.deleteItem()} />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private getLatestItemId(): Promise<number> {

    if (this.props.listTitle.length === 0) {
      this.setState({ status: 'invalid list title' })
    } else {
      const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items?$orderby=Id desc&$top=1&$select=id`;

      return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
        this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => response.json(),
          (error: any): void => reject(error))
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) resolve(-1);
          else resolve(response.value[0].Id);
        });
      });
    }
  }

  private createItem(): void {
    this.setState({ status: 'Creating an item...' });

    const body: string = JSON.stringify({ 'Title': 'new-item' });
    const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items`;

    this.props.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: body
    })
      .then((response: SPHttpClientResponse): Promise<IListItem> => response.json())
      .then((item: IListItem): void =>
        this.setState({ status: `Item "${item.Title}" "${item.Id}" successfully created` }),
        (error: any): void => this.setState({ status: `Error to create an item: ${error}` })
      );
  }

  private readItem(): void {
    this.setState({ status: 'Loading latest item id...' });

    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) throw new Error('No items found in list');
        this.setState({ status: `Loading information about item with id: ${itemId}...` });

        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemId})?$select=Title,Id`;

        return this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => response.json())
      .then((item: IListItem): void =>
        this.setState({ status: `Item Id: ${item.Id}, Title: ${item.Title}` }),
        (error: any): void => this.setState({ status: `Loading latest item failed with error: ${error}` })
      );
  }

  private updateItem(): void {
    this.setState({ status: 'Loading latest item id...' });

    let latestItemId: number;

    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) throw new Error('No items found in list');

        latestItemId = itemId;

        this.setState({ status: `Loading information about item with id: ${itemId}...` });

        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemId})?$select=Title,Id`;

        return this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => response.json())
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        this.setState({ status: 'Pending to update item...' });

        const body: string = JSON.stringify({ 'Title': 'updated-item' });
        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;

        return this.props.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        });
      })
      .then((response: SPHttpClientResponse): void =>
        this.setState({ status: `Item with Id: ${latestItemId} successfully updated` }),
        (error: any) => this.setState({ status: `Error to update an item: ${error}` })
      );
  }

  private deleteItem(): void {
    this.setState({ status: 'Loading latest item id...' });

    let latestItemId: number;
    let etag: string;

    this.getLatestItemId()
      .then((itemId: number): Promise<SPHttpClientResponse> => {
        if (itemId === -1) throw new Error('No items found in the list');

        latestItemId = itemId;

        this.setState({ status: `Loading information about element with id: ${latestItemId}...` });

        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemId})?$select=Id`;
        return this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
        this.setState({ status: `Deleting item with Id: ${item.Id}...` });

        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;

        return this.props.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': etag,
            'X-HTTP-Method': 'DELETE'
          }
        });
      })
      .then((response: SPHttpClientResponse): void =>
        this.setState({ status: `Item with Id: ${latestItemId} successfully deleted` }),
        (error: any) => this.setState({ status: `Error to delete an item: ${error}` })
      );
  }
}
