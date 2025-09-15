import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';

export class DynamicFlowCardPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configure the Dynamic HTML Flow Card" },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Title"
                }),
                PropertyPaneTextField('flowUrl', {
                  label: "Flow URL"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}