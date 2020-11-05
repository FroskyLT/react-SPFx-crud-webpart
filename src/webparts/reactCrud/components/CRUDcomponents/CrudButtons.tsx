import * as React from 'react';
import styles from '../ReactCrud.module.scss';
import { Stack, IStackTokens, DefaultButton, IDropdownOption } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from '../interfaces/IListItem';

interface ICrudButtonsProps {
    siteUrl: string;
    listTitle: string;
    spHttpClient: SPHttpClient;
    itemID: number;
    itemTitle: string;
    addItems(listItems: IDropdownOption[]): void;
    changeStatus(statusText: string): void;
    clearItemTitle(title: string): void;
    changeItemID(id: number): void;
}

interface ICrudButtonsState {
    getHeaders: HeadersInit;
    etag: string;
}
const stackTokens: IStackTokens = { childrenGap: 0 };

export default class ReactCrud extends React.Component<ICrudButtonsProps, ICrudButtonsState> {
    constructor(props: ICrudButtonsProps, state: ICrudButtonsState) {
        super(props);
        this.state = {
            getHeaders: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' },
            etag: ''
        };
    }

    public render(): React.ReactElement<ICrudButtonsProps> {
        return <div>
            <div className={styles.buttonItems}>
                <Stack horizontal tokens={stackTokens}>
                    <DefaultButton text={"Create"} onClick={() => this.createItem()} />
                    <DefaultButton text={"Read"} onClick={() => this.readItem()} />
                    <DefaultButton text={"Update"} onClick={() => this.updateItem()} />
                    <DefaultButton text={"Delete"} onClick={() => this.deleteItem()} />
                </Stack>
            </div>
            <div className={styles.buttonItems}>
                <DefaultButton text={"Get all items"} onClick={() => this.getAll()} />
            </div>
        </div>;
    }

    private async getLatestItem(): Promise<IListItem> {
        try {
            //const itemID: number = await this.getLatestItemId();
            const itemID: number = this.props.itemID;

            if (itemID === 0) throw new Error('Choose the item first');

            this.props.changeStatus(`Loading information about item with id: ${itemID}...`);
            const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${itemID})?$select=Title,Id`;

            const response: SPHttpClientResponse = await this.props.spHttpClient
                .get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
            const item: IListItem = await response.json();

            if (typeof item === 'undefined') throw new Error('Choose the item first');

            this.setState({ etag: response.headers.get('ETag') });
            return item;

        } catch (error) {
            this.props.changeStatus(`Error to read an item: ${error.message}`);
        }
    }

    private async createItem(): Promise<SPHttpClientResponse> {

        this.props.changeStatus('Creating an item...');

        let newTitle: string = 'new-item';
        if (this.props.itemTitle.length !== 0) {
            newTitle = this.props.itemTitle;
            this.props.clearItemTitle("");
        }
        const body: string = JSON.stringify({ 'Title': newTitle });
        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items`;
        const createHeaders: HeadersInit =
            { 'Content-type': 'application/json;odata=nometadata', ...this.state.getHeaders };

        try {
            const response: SPHttpClientResponse = await this.props.spHttpClient
                .post(listUrl, SPHttpClient.configurations.v1, { headers: createHeaders, body: body });
            const item: IListItem = await response.json();

            if (typeof item.Id === 'undefined') throw new Error('there is no such list');

            //this.props.changeStatus(`Item "${item.Title}" "${item.Id}" successfully created`);
            this.getAll(`Item with title: "${item.Title}" and id: "${item.Id}" successfully created`);


            return response;
        } catch (error) {
            this.props.changeStatus(`Error to create an item: ${error.message}`);
        }
    }

    private readItem(): void {

        this.props.changeStatus('Loading choosed item...');

        this.getLatestItem().then((item: IListItem) =>
            this.props.changeStatus(`Item Id: ${item.Id}, Title: ${item.Title}`));
    }

    private async updateItem(): Promise<SPHttpClientResponse> {

        this.props.changeStatus('Loading choosed item...');
        const updateHeaders: HeadersInit = {
            ...this.state.getHeaders,
            'Content-type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
        };

        let newTitle: string = 'updated-item';
        if (this.props.itemTitle.length !== 0) {
            newTitle = this.props.itemTitle;
            this.props.clearItemTitle("");
        }

        try {
            const item: IListItem = await this.getLatestItem();

            if (typeof item === 'undefined') throw new Error('while reading an item');
            this.props.changeStatus('Pending to update item...');

            const body: string = JSON.stringify({ 'Title': newTitle });
            const listUrl: string =
                `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;

            const response: SPHttpClientResponse = await this.props.spHttpClient
                .post(listUrl, SPHttpClient.configurations.v1, { headers: updateHeaders, body: body });

            //this.props.changeStatus(`Item with Id: ${item.Id} successfully updated`);
            this.getAll(`Item with Id: ${item.Id} successfully updated`);


            return response;
        } catch (error) {
            this.props.changeStatus(`Error to update an item: ${error.message}`);
        }
    }

    private async deleteItem(): Promise<SPHttpClientResponse> {

        this.props.changeStatus('Loading choosed item...');

        try {
            const item: IListItem = await this.getLatestItem();
            const deleteHeaders: HeadersInit = {
                ...this.state.getHeaders,
                'Content-type': 'application/json;odata=nometadata',
                'IF-MATCH': this.state.etag,
                'X-HTTP-Method': 'DELETE'
            };

            if (typeof item === 'undefined') throw new Error('while reading an item');
            this.props.changeStatus(`Deleting item with Id: ${item.Id}...`);

            const listUrl: string =
                `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items(${item.Id})`;

            const response: SPHttpClientResponse = await this.props.spHttpClient
                .post(listUrl, SPHttpClient.configurations.v1, { headers: deleteHeaders });

            this.props.changeItemID(0);
            //this.props.changeStatus(`Item with Id: ${item.Id} successfully deleted`);
            this.getAll(`Item with Id: ${item.Id} successfully deleted`);


            return response;

        } catch (error) {
            this.props.changeStatus(`Error to delete an item: ${error.message}`);
        }
    }


    private async getAll(success: string = 'All items were read successfully'): Promise<SPHttpClientResponse> {

        const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items?$select=Title,Id`;
        const listItems: IDropdownOption[] = [];

        try {
            const response: SPHttpClientResponse = await this.props.spHttpClient
                .get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
            const items: { value: IListItem[] } = await response.json();

            if (typeof items.value === 'undefined') {
                this.props.addItems(listItems);
                throw new Error('there is no such list');
            }
            else if (items.value.length === 0) {
                this.props.addItems(listItems);
                throw new Error('there is no items in the list');
            }
            else {
                this.props.changeStatus(success);
                items.value.map((i) => listItems.push({ key: i.Id, text: i.Title }));
                this.props.addItems(listItems);
            }

            return response;
        } catch (error) {
            this.props.changeStatus(`Error to read all items: ${error.message}`);
        }
    }
}



// private async getLatestItemId(): Promise<number> {

//     const listUrl: string = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.listTitle}')/items?$orderby=Id desc&$top=1&$select=id`;

//     const response: SPHttpClientResponse = await this.props.spHttpClient
//         .get(listUrl, SPHttpClient.configurations.v1, { headers: this.state.getHeaders });
//     const data: { value: { Id: number }[] } = await response.json();

//     if (typeof data === 'undefined') throw new Error('there is no such list');
//     if (data.value.length === 0) return -1;
//     else return data.value[0].Id;

// }