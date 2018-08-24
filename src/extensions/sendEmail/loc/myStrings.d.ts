declare interface ISendEmailCommandSetStrings {
  spfxEmailTo: string;
}

declare module 'SendEmailCommandSetStrings' {
  const strings: ISendEmailCommandSetStrings;
  export = strings;
}
