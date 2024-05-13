declare interface IBulkUpdateStrings {
  Command1: string;
  Command2: string;
}

declare module 'BulkUpdateStrings' {
  const strings: IBulkUpdateStrings;
  export = strings;
}
