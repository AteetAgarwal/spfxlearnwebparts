import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SourceconnectecwebpartWebPartStrings';
import Sourceconnectecwebpart from './components/Sourceconnectecwebpart';
import { ISourceconnectecwebpartProps } from './components/ISourceconnectecwebpartProps';
import {IDynamicDataCallables, IDynamicDataPropertyDefinition} from '@microsoft/sp-dynamic-data';

export interface ISourceconnectecwebpartWebPartProps {
  description: string;
}

export interface IList{
  Title:string;
}

export default class SourceconnectecwebpartWebPart extends BaseClientSideWebPart<ISourceconnectecwebpartWebPartProps> implements IDynamicDataCallables {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  public ListTitle:IList;

  public render(): void {
    const element: React.ReactElement<ISourceconnectecwebpartProps> = React.createElement(
      Sourceconnectecwebpart,
      {
        description: this.properties.description,
        context: this.context,
        environmentMessage: this._environmentMessage,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        PassListTitle: this.getTitle
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public getPropertyDefinitions():ReadonlyArray<IDynamicDataPropertyDefinition>{
    return[{id:"Title", title:"Title"}];
  }

  public getPropertyValue(propertyId:string):IList{
    if(propertyId==="Title"){
      return this.ListTitle;
    }
  }

  public getTitle=(title:IList):  void =>{
    this.ListTitle=title;
    this.context.dynamicDataSourceManager.notifyPropertyChanged("Title");
  }

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);
    return new Promise<void>(async (resolve,reject)=>{
      this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
        resolve();
      });
    })
  }


  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
