declare interface IProviderWebPartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'ProviderWebPartWebPartStrings' {
  const strings: IProviderWebPartWebPartStrings;
  export = strings;
}
