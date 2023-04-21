import * as React from 'react';
import styles from './ReactShowListItemsWebPart.module.scss';
import { IReactShowListItemsWebPartProps } from './IReactShowListItemsWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jQuery from 'jquery';

export interface ReactShowListItemsWebPartState {
  listItems: [
    {
      Title: '';
      ID: '';
      SoftwareName: '';
    }
  ];
}

export default class ReactShowListItemsWebPart extends React.Component<
  IReactShowListItemsWebPartProps,
  ReactShowListItemsWebPartState
> {
  static siteUrl: string = '';

  public constructor(
    props: IReactShowListItemsWebPartProps,
    state: ReactShowListItemsWebPartState
  ) {
    super(props);
    this.state = {
      listItems: [
        {
          Title: '',
          ID: '',
          SoftwareName: '',
        },
      ],
    };
    ReactShowListItemsWebPart.siteUrl = this.props.websiteUrl;
  }

  public componentDidMount() {
    let reactContextHandler = this;

    jQuery.ajax({
      url: `${ReactShowListItemsWebPart.siteUrl}/_api/web/lists/getByTitle('MicrosoftSoftware')/items`,
      type: 'GET',
      headers: { Accept: 'application/json; odata=verbose;' },
      success: function (resultData) {
        reactContextHandler.setState({
          listItems: resultData.d.results,
        });
      },
      error: function (jqXHR, textStatus, errorThrown) {},
    });
  }

  public render(): React.ReactElement<IReactShowListItemsWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.reactShowListItemsWebPart} ${
          hasTeamsContext ? styles.teams : ''
        }`}>
        <div className={styles.rShowListItems}></div>
      </section>
    );
  }
}
