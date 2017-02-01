declare interface ISpFxFlowEmailStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  EmailGroupName: string;
  FlowURLLabel: string;
  EmailAddressFieldLabel: string;
  EmailSubjectFieldLabel: string;
  EmailBodyFieldLabel: string;
}

declare module 'spFxFlowEmailStrings' {
  const strings: ISpFxFlowEmailStrings;
  export = strings;
}
