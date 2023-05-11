declare interface IElevatepermissionWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  FlowUrlFieldLabel:string;
  ListTitleFieldLabel:string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'ElevatepermissionWebPartStrings' {
  const strings: IElevatepermissionWebPartStrings;
  export = strings;
}
