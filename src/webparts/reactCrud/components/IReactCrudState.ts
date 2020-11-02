import { IListItem } from './IListItem';
export interface IReactCrudState {
  status: string;
  getHeaders: HeadersInit;
  items?: IListItem[];
}
