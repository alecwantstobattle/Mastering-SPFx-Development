import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './FieldExtensionReact.module.scss';

export interface IFieldExtensionReactProps {
  text: string;
}

const LOG_SOURCE: string = 'FieldExtensionReact';

export default class FieldExtensionReact extends React.Component<
  IFieldExtensionReactProps,
  {}
> {
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: FieldExtensionReact mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: FieldExtensionReact unmounted');
  }

  public render(): React.ReactElement<{}> {
    const myStyles = {
      color: 'blue',
      width: `${this.props.text}px`,
      background: 'red',
    };

    return (
      <div className={styles.FieldExtensionReact}>
        <div style={myStyles}>{this.props.text}</div>
      </div>
    );
  }
}
