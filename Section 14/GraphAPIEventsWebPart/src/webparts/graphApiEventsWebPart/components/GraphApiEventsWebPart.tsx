import * as React from 'react';
import styles from './GraphApiEventsWebPart.module.scss';
import { IGraphApiEventsWebPartProps } from './IGraphApiEventsWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IGraphApiEventsState } from './IGraphApiEventsState';

export default class GraphApiEventsWebPart extends React.Component<
  IGraphApiEventsWebPartProps,
  IGraphApiEventsState
> {
  constructor(props: IGraphApiEventsWebPartProps) {
    super(props);
    this.state = {
      events: [],
    };
  }

  public componentDidMount(): void {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/me/calendar/events')
          .version('v1.0')
          .select('*')
          .get((error: any, eventsResponse, rawResponse?: any) => {
            if (error) {
              console.error('Message is : ' + error);
              return;
            }

            const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
            this.setState({ events: calendarEvents });
          });
      });
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
            {this.state.events.map((item, key) => (
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
