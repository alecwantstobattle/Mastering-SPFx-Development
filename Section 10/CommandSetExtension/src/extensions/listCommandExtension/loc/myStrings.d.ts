declare interface IListCommandExtensionCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ListCommandExtensionCommandSetStrings' {
  const strings: IListCommandExtensionCommandSetStrings;
  export = strings;
}
