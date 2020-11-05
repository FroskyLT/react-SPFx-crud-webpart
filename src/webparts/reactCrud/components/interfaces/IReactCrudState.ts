import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
export interface IReactCrudState {
  listTitle: string;
  itemTitle?: string;
  items?: IDropdownOption[];
  itemID?: number;
  status: string;
}
