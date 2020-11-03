import { IListItem } from './IListItem';
export interface IReactCrudState {
  listTitle: string;
  status: string;
  items?: IListItem[];
}
