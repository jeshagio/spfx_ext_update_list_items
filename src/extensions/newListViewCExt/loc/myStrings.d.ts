declare interface INewListViewCExtCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'NewListViewCExtCommandSetStrings' {
  const strings: INewListViewCExtCommandSetStrings;
  export = strings;
}
