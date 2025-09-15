import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';

export class AceDynamicFlowCardPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configure the Dynamic HTML Flow Card to display HTML content from a Power Automate endpoint." },
          groups: [
            {
              groupName: "Basic Settings",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Card title",
                  description: "Title displayed on the ACE card"
                }),
                PropertyPaneTextField('buttonLabel', {
                  label: "Button label",
                  description: "Text displayed on the card button (e.g., 'View Details', 'Get Report', 'Load Content')",
                  value: "View Content" // Default value
                }),
                PropertyPaneTextField('flowUrl', {
                  label: "Flow URL",
                  description: "Enter the URL of the flow that returns HTML content",
                  multiline: false
                })
              ]
            },
            {
              groupName: "Content Settings",
              groupFields: [
                PropertyPaneTextField('prompt', {
                  label: "Prompt",
                  description: "Enter a prompt or instruction to send to the Power Automate flow",
                  multiline: true,
                  rows: 4
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
