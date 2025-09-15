import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../AceDynamicFlowCardAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState
> {
  /**
   * Buttons will be visible for both 'Medium' and 'Large' card sizes with Basic Card View.
   * It will support up to two buttons.
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    const buttonLabel = this.properties.buttonLabel || 'View Content'; // Fallback to default
    
    return [
      {
        title: buttonLabel,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IBasicCardParameters {
    return {
      primaryText: this.properties.title || 'Dynamic Flow Card',
      title: this.properties.title || 'Dynamic Flow Card'
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }
}
