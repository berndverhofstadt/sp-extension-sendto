declare interface ISendEmailCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'SendEmailCommandSetStrings' {
  const strings: ISendEmailCommandSetStrings;
  export = strings;
}
