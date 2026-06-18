import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { CSS, HTML } from "./markup";
import { initCalculator } from "./ui";

/**
 * Division 296 reset calculator — SPFx web part wrapper.
 *
 * Injects the verified site markup + styles (generated from web/ by
 * build-spfx-markup.mjs) into the web part's own DOM element and boots the
 * client-side calculator. No properties, no API calls, no external requests —
 * everything runs in the browser. Rendered once so user input survives
 * incidental re-renders (resize / property-pane).
 */
export interface IDiv296CalculatorWebPartProps {}

export default class Div296CalculatorWebPart extends BaseClientSideWebPart<IDiv296CalculatorWebPartProps> {
  private _booted = false;

  public render(): void {
    if (this._booted) {
      return;
    }
    this.domElement.innerHTML = `<style>${CSS}</style><div class="div296-root">${HTML}</div>`;
    initCalculator(this.domElement);
    this._booted = true;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
