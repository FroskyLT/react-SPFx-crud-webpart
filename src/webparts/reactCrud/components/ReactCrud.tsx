import * as React from 'react';
import styles from './ReactCrud.module.scss';
import { IReactCrudProps } from './IReactCrudProps';
import { IReactCrudState } from './IReactCrudState';
import { escape } from '@microsoft/sp-lodash-subset';
import ListTitleTextfield from './ListTitleTextfield';
import CrudButtons from './CrudButtons';

export default class ReactCrud extends React.Component<IReactCrudProps, IReactCrudState> {
  constructor(props: IReactCrudProps, state: IReactCrudState) {
    super(props);
    this.state = {
      listTitle: "",
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
              <ListTitleTextfield handleSubmit={this.handleSubmit} />
              <p className={styles.description}>Status is: {this.state.status}</p>
            </div>
            <div className={styles.column}>
              <CrudButtons siteUrl={this.props.siteUrl} listTitle={this.state.listTitle}
                spHttpClient={this.props.spHttpClient} changeStatus={this.changeStatus} />
            </div>
          </div>
        </div>
      </div>
    );
  }

  private handleSubmit = (title: string): void => { this.setState({ listTitle: title }); };
  private changeStatus = (statusText: string): void => { this.setState({ status: statusText }); };

}
