declare interface ISpfxlearnwpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ListTitleFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  ListUrlFieldLabel: string;
  PercentCompletedFieldLabel: string;
  ValidationRequiredFieldLabel: string;
  ListNameFieldLabel:string;
}

declare module 'SpfxlearnwpWebPartStrings' {
  const strings: ISpfxlearnwpWebPartStrings;
  export = strings;
}
