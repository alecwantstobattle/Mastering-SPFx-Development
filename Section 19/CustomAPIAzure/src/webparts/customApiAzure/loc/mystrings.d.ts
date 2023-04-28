declare interface ICustomApiAzureWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'CustomApiAzureWebPartStrings' {
  const strings: ICustomApiAzureWebPartStrings;
  export = strings;
}
