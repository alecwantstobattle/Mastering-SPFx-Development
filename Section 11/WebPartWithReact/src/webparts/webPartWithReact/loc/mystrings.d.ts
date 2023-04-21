declare interface IWebPartWithReactWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'WebPartWithReactWebPartStrings' {
  const strings: IWebPartWithReactWebPartStrings;
  export = strings;
}
