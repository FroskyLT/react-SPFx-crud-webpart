import * as React from 'react';
import styles from './ReactCrud.module.scss';
import { PrimaryButton } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from './IListItem';

interface ICrudButtonsProps {
    siteUrl: string;
    listTitle: string;
    spHttpClient: SPHttpClient;
    changeStatus(statusText: string): void;
}

interface ICrudButtonsState {
    getHeaders: HeadersInit;
}

export default class ReactCrud extends React.Component<ICrudButtonsProps, ICrudButtonsState> {
    constructor(props: ICrudButtonsProps, state: ICrudButtonsState) {
        super(props);
        this.state = {
            getHeaders: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
        };
    }

    public render(): React.ReactElement<ICrudButtonsProps> {
        return <div className={styles.buttonItems}>
            <PrimaryButton text={"Create"} onClick={() => this.createItem()} />
            <PrimaryButton text={"Read"} onClick={() => this.readItem()} />
            <PrimaryButton text={"Update"} onClick={() => this.updateItem()} />
            <PrimaryButton text={"Delete"} onClick={() => this.deleteItem()} />
        </div>;
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
            this.props.changeStatus(`Loading information about item with id: ${itemID}...`);

            const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemID})?$select=Title,Id`;

            const response: SPHttpClientResponse = await this.props.spHttpClient
                .get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
            const item: IListItem = await response.json();


            return item;

        } catch (error) {
            this.props.changeStatus(`Error to read an item: ${error.message}`);
        }
    }


    private createItem(): void {

        this.props.changeStatus('Creating an item...');

        const body: string = JSON.stringify({ 'Title': 'new-item' });
        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items`;
        const createHeaders: HeadersInit =
            { 'Content-type': 'application/json;odata=nometadata', ...this.state.getHeaders };

        const postItem: () => Promise<IListItem> = async () => {
            try {
                const response: SPHttpClientResponse = await this.props.spHttpClient
                    .post(listUrl, SPHttpClient.configurations.v1, { headers: createHeaders, body: body });
                const item: IListItem = await response.json();

                if (typeof item.Id === 'undefined') throw new Error('Something went wrong');

                return item;
            } catch (error) {
                this.props.changeStatus(`Error to create an item: ${error.message}`);
            }
        };

        postItem().then((item: IListItem): void =>
            this.props.changeStatus(`Item "${item.Title}" "${item.Id}" successfully created`));
    }

    private readItem(): void {

        this.props.changeStatus('Loading latest item...');

        this.getLatestItem().then((item: IListItem) =>
            this.props.changeStatus(`Item Id: ${item.Id}, Title: ${item.Title}`));
    }

    private updateItem(): void {

        this.props.changeStatus('Loading latest item...');

        let latestItemId: number;
        const updateHeaders: HeadersInit =
        {
            ...this.state.getHeaders,
            'Content-type': 'application/json;odata=nometadata',
            'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE'
        };

        const update: () => void = async () => {
            try {
                const item: IListItem = await this.getLatestItem();

                if (typeof item === 'undefined') throw new Error('while reading an item');
                latestItemId = item.Id;
                this.props.changeStatus('Pending to update item...');
                const body: string = JSON.stringify({ 'Title': 'updated-item' });
                const listUrl: string =
                    `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;

                const response: SPHttpClientResponse = await this.props.spHttpClient
                    .post(listUrl, SPHttpClient.configurations.v1, { headers: updateHeaders, body: body });

                this.props.changeStatus(`Item with Id: ${latestItemId} successfully updated`);


            } catch (error) {
                this.props.changeStatus(`Error to update an item: ${error.message}`);
            }
        };

        update();
    }

    private deleteItem(): void {
        this.props.changeStatus('Loading latest item...');

        let latestItemId: number;
        let etag: string;

        const getLItem: () => Promise<IListItem> = async () => {
            try {
                const itemID: number = await this.getLatestItemId();

                if (itemID === -1) throw new Error('No items found in list');
                this.props.changeStatus(`Loading information about item with id: ${itemID}...`);
                const listUrl: string =
                    `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemID})?$select=Title,Id`;

                const response: SPHttpClientResponse = await this.props.spHttpClient
                    .get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });

                const item: IListItem = await response.json();


                etag = response.headers.get('ETag');
                latestItemId = item.Id;

                return item;

            } catch (error) {
                this.props.changeStatus(`Error to read an item: ${error}`);
            }
        };

        const deleteLItem: () => void = async () => {
            try {
                const item: IListItem = await getLItem();
                if (typeof item === 'undefined') throw new Error('while reading an item');

                this.props.changeStatus(`Deleting item with Id: ${item.Id}...`);
                const listUrl: string =
                    `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;
                const deleteHeaders: HeadersInit =
                {
                    ...this.state.getHeaders,
                    'Content-type': 'application/json;odata=nometadata',
                    'IF-MATCH': etag, 'X-HTTP-Method': 'DELETE'
                };

                const response: SPHttpClientResponse = await this.props.spHttpClient
                    .post(listUrl, SPHttpClient.configurations.v1, { headers: deleteHeaders });


                this.props.changeStatus(`Item with Id: ${latestItemId} successfully deleted`);

            } catch (error) {
                this.props.changeStatus(`Error to delete an item: ${error.message}`);
            }
        };
        deleteLItem();
    }
}
