import { SPHttpClient } from '@microsoft/sp-http';
export interface IReactCrudProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
