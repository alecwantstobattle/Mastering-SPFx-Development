import * as React from 'react';
import styles from './ReactLifeCycleWebPart.module.scss';
import { IReactLifeCycleWebPartProps } from './IReactLifeCycleWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IReactLifeCycleWebPartState {
  stageTitle: string;
}

export default class ReactLifeCycleWebPart extends React.Component<
  IReactLifeCycleWebPartProps,
  IReactLifeCycleWebPartState
> {
  constructor(
    props: IReactLifeCycleWebPartProps,
    state: IReactLifeCycleWebPartState
  ) {
    super(props);

    this.state = {
      stageTitle: 'component Constructor has been called.',
    };

    this.updateState = this.updateState.bind(this);

    console.log('Stage Title from Constructor: ' + this.state.stageTitle);
  }

  componentWillMount(): void {
    console.log('Component will mount has been called.');
  }

  componentDidMount(): void {
    console.log('Stage Title from componentDidMount: ' + this.state.stageTitle);
    this.setState({
      stageTitle: 'componentDidMount has been called.',
    });
  }

  public updateState() {
    this.setState({
      stageTitle: 'changeState has been called.',
    });
  }

  public render(): React.ReactElement<IReactLifeCycleWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.reactLifeCycleWebPart} ${
          hasTeamsContext ? styles.teams : ''
        }`}>
        <div>
          <h1>ReactJS component Lifecycle</h1>
          <h3>{this.state.stageTitle}</h3>
          <button onClick={this.updateState}>
            Click here to Update State Data!
          </button>
        </div>
      </section>
    );
  }

  componentWillUnmount(): void {
    console.log('Component will unmount has been called.');
  }
}
