import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';

import * as strings from 'SpBannerApplicationCustomizerStrings';
const LOG_SOURCE: string = 'SpBannerApplicationCustomizer';

// Reference to the 2013 solution external library
import "sp-banner-2013";
declare var SPBanner;

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpBannerApplicationCustomizerProperties { }

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpBannerApplicationCustomizer
  extends BaseApplicationCustomizer<ISpBannerApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Log
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Handle possible changes on the existence of placeholders
    this.context.placeholderProvider.changedEvent.add(this, this.renderBanner);

    // Resolve the promise
    return Promise.resolve();
  }

  // Renders the banner
  private _elBanner: HTMLElement = null;
  private renderBanner() {
    // Do nothing if we have already created the banner
    if (this._elBanner) { return; }

    // Create the element
    this._elBanner = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top).domElement;

    // Generate the banner
    SPBanner(this._elBanner);
  }
}
