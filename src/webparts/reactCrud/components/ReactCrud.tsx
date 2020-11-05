import * as React from 'react';
import styles from './ReactCrud.module.scss';
import { IDropdownOption } from 'office-ui-fabric-react';
import { escape } from '@microsoft/sp-lodash-subset';
import { IReactCrudProps } from './interfaces/IReactCrudProps';
import { IReactCrudState } from './interfaces/IReactCrudState';

import ListTitleTextfield from './CRUDcomponents/ListTitleTextfield';
import NewItemTitleTextField from './CRUDcomponents/NewItemTitleTextField';
import CrudButtons from './CRUDcomponents/CrudButtons';
import DropdownItems from './CRUDcomponents/DropdownItems';

export default class ReactCrud extends React.Component<IReactCrudProps, IReactCrudState> {
  constructor(props: IReactCrudProps, state: IReactCrudState) {
    super(props);
    this.state = {
      listTitle: '',
      itemTitle: '',
      items: [],
      itemID: 0,
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
              <p className={styles.subTitle}>Your Sharepoint list title: {escape(this.state.listTitle)}</p>
              <ListTitleTextfield changeListTitle={this.handleListTitle} />
            </div>

            <div className={styles.column}>
              <div className={styles.customColumn}>
                <div className={styles.customColumn__item}>
                  <DropdownItems
                    items={this.state.items} chooseItem={this.handleSingleItem} />
                </div>
                <div className={styles.customColumn__item}>
                  <NewItemTitleTextField
                    title={this.state.itemTitle} changeItemTitle={this.handleItemTitle} />
                </div>
              </div>
            </div>

            <div className={styles.column}>
              <CrudButtons
                siteUrl={this.props.siteUrl}
                spHttpClient={this.props.spHttpClient}
                listTitle={this.state.listTitle}
                itemTitle={this.state.itemTitle}
                itemID={this.state.itemID}
                changeStatus={this.handleStatus}
                clearItemTitle={this.handleItemTitle}
                addItems={this.handleItems}
                changeItemID={this.handleItemID} />
              <p className={styles.description}>Status is: {this.state.status}</p>
            </div>

          </div>
        </div>
      </div>
    );
  }

  private handleListTitle = (title: string): void => { this.setState({ listTitle: title }); };
  private handleItemTitle = (title: string): void => { this.setState({ itemTitle: title }); };
  private handleStatus = (statusText: string): void => { this.setState({ status: statusText }); };
  private handleItemID = (id: number): void => { this.setState({ itemID: id }); };

  private handleItems = (listItems: IDropdownOption[]): void => {
    if (listItems.length === 0) this.setState({ items: [] });
    else this.setState({ items: listItems });
  }
  private handleSingleItem = (item: any): void => {
    this.setState({ itemID: item.key, status: `choosed item with Id: ${item.key}, Title: ${item.text}` });
  }

}
