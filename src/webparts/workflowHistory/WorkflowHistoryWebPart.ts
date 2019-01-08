import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import { setup as pnpSetup } from "@pnp/common";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'WorkflowHistoryWebPartStrings';
import WorkflowHistory from './components/WorkflowHistory';
import { IWorkflowHistoryProps } from './components/IWorkflowHistoryProps';

export interface IWorkflowHistoryWebPartProps {
  description: string;
}

export default class WorkflowHistoryWebPart extends BaseClientSideWebPart<IWorkflowHistoryWebPartProps> {
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      pnpSetup({
        spfxContext: this.context
      });
    });
  }
  
  public render(): void {
    const element: React.ReactElement<IWorkflowHistoryProps > = React.createElement(
      WorkflowHistory,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
