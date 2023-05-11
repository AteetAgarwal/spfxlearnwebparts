import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxlearnwpWebPart.module.scss';
import * as strings from 'SpfxlearnwpWebPartStrings';
import * as $ from 'jquery';
import "jqueryui";
import {SPComponentLoader} from '@microsoft/sp-loader'; 

export interface ISpfxlearnwpWebPartProps {
  ListTitle: string;
  ListUrl:string;
  PercentCompleted: number;
  ValidationRequired: boolean;
  ListName: string;
}

export default class SpfxlearnwpWebPart extends BaseClientSideWebPart<ISpfxlearnwpWebPartProps> {
  public constructor(){
    super();
    SPComponentLoader.loadCss("//code.jquery.com/ui/1.11.1/themes/smoothness/jquery-ui.css");
  }
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
      <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
      <div class="${styles.spfxlearnwp}">${this._environmentMessage}</div>
      <div class="accordion">
        <h3>Step1</h3>
        <div>
          <p>
            Demo jquery and jquery ui
          </p>
        </div>
        <h3>Step2</h3>
        <div>
          <p>
            Accordion demo
          </p>
        </div>
    </div>
    `;
    const accOptions : JQueryUI.AccordionOptions={
      animate: true,
      collapsible: true
    }
    $(".accordion").accordion(accOptions);    
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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

  public ValidateListUrl(value:string):string{
    if(value.length>255){
      return "URL should be less than 256 characters";
    }
    if(value.length==0){
      return "Enter the valid list url";
    }
    return "";
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
                PropertyPaneTextField('ListTitle', {
                  label: strings.ListTitleFieldLabel
                }),
                PropertyPaneTextField('ListUrl', {
                  label: strings.ListUrlFieldLabel,
                  onGetErrorMessage: this.ValidateListUrl.bind(this)
                }),
                PropertyPaneSlider('PercentCompleted',{
                  label: strings.PercentCompletedFieldLabel,
                  min: 0,
                  max:100
                }),
                PropertyPaneCheckbox('ValidationRequired',{
                  text: strings.ValidationRequiredFieldLabel
                }),
                PropertyPaneDropdown('ListName',{
                  label: strings.ListNameFieldLabel,
                  options:[{
                    key:"--Select your list--",
                    text:"--Select your list--"
                  },
                  {
                    key:"Documents",
                    text:"Documents"
                  }],
                  selectedKey: "--Select your list--"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
