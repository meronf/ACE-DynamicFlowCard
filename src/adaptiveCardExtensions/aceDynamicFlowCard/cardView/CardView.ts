import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../AceDynamicFlowCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IAceDynamicFlowCardAdaptiveCardExtensionProps,
  IAceDynamicFlowCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: "Dynamic HTML Content"
      },
      footer: {
        componentName: 'cardButton',
        title: "View HTML",
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    });
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
