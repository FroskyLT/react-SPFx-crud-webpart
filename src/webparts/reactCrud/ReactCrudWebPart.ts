import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactCrudWebPartStrings';
import ReactCrud from './components/ReactCrud';
import { IReactCrudProps } from './components/interfaces/IReactCrudProps';

export interface IReactCrudWebPartProps {
  listTitle: string;
}

export default class ReactCrudWebPart extends BaseClientSideWebPart<IReactCrudWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactCrudProps> = React.createElement(
      ReactCrud,
      {
        //listTitle: this.properties.listTitle,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
