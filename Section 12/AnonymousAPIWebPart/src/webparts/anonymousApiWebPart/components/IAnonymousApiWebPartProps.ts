import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAnonymousApiWebPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  apiUrl: string;
  userId: string;
  context: WebPartContext;
}
