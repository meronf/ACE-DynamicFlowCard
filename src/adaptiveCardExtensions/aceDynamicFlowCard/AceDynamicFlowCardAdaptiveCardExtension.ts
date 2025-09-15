import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { AceDynamicFlowCardPropertyPane } from './AceDynamicFlowCardPropertyPane';

export interface IAceDynamicFlowCardAdaptiveCardExtensionProps {
  title: string;
  flowUrl: string; // Make sure this matches your property pane
  prompt: string; // New prompt property
  powerAutomateUrl?: string; // Keep for backward compatibility
}

export interface IAceDynamicFlowCardAdaptiveCardExtensionState {
  htmlContent: string;
  isLoading: boolean;
  error: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'AceDynamicFlowCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'AceDynamicFlowCard_QUICK_VIEW';

export default class AceDynamicFlowCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AceDynamicFlowCardPropertyPane;

  public onInit(): Promise<void> {
    this.state = { 
      htmlContent: '',
      isLoading: true,
      error: ''
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'AceDynamicFlowCard-property-pane'*/
      './AceDynamicFlowCardPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.AceDynamicFlowCardPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
