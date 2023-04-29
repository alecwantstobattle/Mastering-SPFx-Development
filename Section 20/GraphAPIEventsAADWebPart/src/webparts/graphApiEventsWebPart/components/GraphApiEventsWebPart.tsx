import * as React from 'react';
import styles from './GraphApiEventsWebPart.module.scss';
import { IGraphApiEventsWebPartProps } from './IGraphApiEventsWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  AadHttpClient,
  AadHttpClientFactory,
  HttpClientResponse,
} from '@microsoft/sp-http';

import { IGraphApiEventsState } from './IGraphApiEventsState';

export default class GraphApiEventsWebPart extends React.Component<
  IGraphApiEventsWebPartProps,
  IGraphApiEventsState
> {
  constructor(props: IGraphApiEventsWebPartProps) {
    super(props);
    this.state = {
      events: [],
      events2: [],
    };
  }

  public getCalendarEventsUsingGraphClient(): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
            .api('/me/calendar/events')
            .version('v1.0')
            .select('*')
            .get((error: any, eventsResponse, rawResponse?: any) => {
              resolve(eventsResponse.value);
            })
            .catch((error) => {
              reject(error);
            });
        });
    });
  }

  public getCalendarEventsUsingAadHttpClient(): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      this.props.context.aadHttpClientFactory
        .getClient('https://graph.microsoft.com')
        .then((client: AadHttpClient) => {
          client
            .get(
              'https://graph.microsoft.com/v1.0/me/calendar/events',
              AadHttpClient.configurations.v1
            )
            .then((response: HttpClientResponse) => {
              return response.json();
            })
            .then((jsonResponse: any) => {
              resolve(jsonResponse);
            })
            .catch((error) => {
              reject(error);
            });
        });
    });
  }

  public componentDidMount(): void {
    this.getCalendarEventsUsingGraphClient().then((allCalendarEvents: any) => {
      this.setState({ events2: allCalendarEvents });
    });

    this.getCalendarEventsUsingAadHttpClient().then(
      (allCalendarEvents: any) => {
        this.setState({ events: allCalendarEvents.value });
      }
    );
  }

  public render(): React.ReactElement<IGraphApiEventsWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.graphApiEventsWebPart} ${
          hasTeamsContext ? styles.teams : ''
        }`}>
        <div>
          <ul>
            {this.state.events.map((item, key) => (
              <li key={item.id}>
                {item.subject},{item.organizer.emailAddress.name},
                {item.start.dateTime.substr(0, 10)},
                {item.start.dateTime.substr(12, 5)},
                {item.end.dateTime.substr(0, 10)},
                {item.end.dateTime.substr(12, 5)}
              </li>
            ))}
          </ul>
          <style>{`
table {
border:1px solid black;
background-color:aqua;

} `}</style>
          <table>
            <tr>
              <td>Subject</td>
              <td>Organizer Name</td>
              <td>Start Date</td>
              <td>Start Time</td>
              <td>End Date</td>
              <td>End Time</td>
            </tr>
            {this.state.events2.map((item, key) => (
              <tr>
                <td>{item.subject}</td>
                <td>{item.organizer.emailAddress.name}</td>
                <td>{item.start.dateTime.substr(0, 10)}</td>
                <td>{item.start.dateTime.substr(12, 5)}</td>
                <td>{item.end.dateTime.substr(0, 10)}</td>
                <td>{item.end.dateTime.substr(12, 5)}</td>
              </tr>
            ))}
          </table>
        </div>
      </section>
    );
  }
}
