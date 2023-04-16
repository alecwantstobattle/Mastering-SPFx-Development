import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters,
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'FieldCustomizerExtensionFieldCustomizerStrings';
import styles from './FieldCustomizerExtensionFieldCustomizer.module.scss';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFieldCustomizerExtensionFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'FieldCustomizerExtensionFieldCustomizer';

export default class FieldCustomizerExtensionFieldCustomizer extends BaseFieldCustomizer<IFieldCustomizerExtensionFieldCustomizerProperties> {
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(
      LOG_SOURCE,
      'Activated FieldCustomizerExtensionFieldCustomizer with properties:'
    );
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(
      LOG_SOURCE,
      `The following string should be equal: "FieldCustomizerExtensionFieldCustomizer" and "${strings.Title}"`
    );
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const text: string = `${event.fieldValue}`;

    event.domElement.innerHTML = `<div class='${styles.FieldCustomizerExtension}'>
      <div class='${styles.cell}'>
        <div style='background:red;width:${event.fieldValue}px;color:blue;'>${event.fieldValue}%</div>
      </div>
    </div>`;
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}
