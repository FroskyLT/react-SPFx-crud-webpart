import {SPHttpClient} from '@microsoft/sp-http';
export interface IReactCrudProps {
  listTitle: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
