declare interface IHelloWorld2WebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'HelloWorld2WebPartStrings' {
  const strings: IHelloWorld2WebPartStrings;
  export = strings;
}
