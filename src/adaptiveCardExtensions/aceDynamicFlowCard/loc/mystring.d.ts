declare interface IAceDynamicFlowCardAdaptiveCardExtensionStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  PowerAutomateUrlFieldLabel: string;
  PowerAutomateUrlFieldDescription: string;
  FlowUrlFieldLabel: string; // Add this new property
  FlowUrlFieldDescription: string; // Add this new property
  Title: string;
  SubTitle: string;
  PrimaryText: string;
  Description: string;
  QuickViewButton: string;
  LoadingMessage: string;
  ErrorMessage: string;
}

declare module 'AceDynamicFlowCardAdaptiveCardExtensionStrings' {
  const strings: IAceDynamicFlowCardAdaptiveCardExtensionStrings;
  export = strings;
}
