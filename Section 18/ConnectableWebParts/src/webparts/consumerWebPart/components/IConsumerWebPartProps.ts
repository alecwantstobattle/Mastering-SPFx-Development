import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IConsumerWebPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  siteUrl: string;
  DeptTitleId: DynamicProperty<string>;
}
