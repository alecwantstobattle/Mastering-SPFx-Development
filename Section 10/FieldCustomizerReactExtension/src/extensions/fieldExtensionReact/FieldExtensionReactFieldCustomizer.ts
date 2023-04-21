import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters,
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'FieldExtensionReactFieldCustomizerStrings';
import FieldExtensionReact, {
  IFieldExtensionReactProps,
} from './components/FieldExtensionReact';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFieldExtensionReactFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'FieldExtensionReactFieldCustomizer';

export default class FieldExtensionReactFieldCustomizer extends BaseFieldCustomizer<IFieldExtensionReactFieldCustomizerProperties> {
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(
      LOG_SOURCE,
      'Activated FieldExtensionReactFieldCustomizer with properties:'
    );
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(
      LOG_SOURCE,
      `The following string should be equal: "FieldExtensionReactFieldCustomizer" and "${strings.Title}"`
    );
    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const text: string = `${event.fieldValue}`;

    const fieldExtensionReact: React.ReactElement<{}> = React.createElement(
      FieldExtensionReact,
      { text } as IFieldExtensionReactProps
    );

    ReactDOM.render(fieldExtensionReact, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
