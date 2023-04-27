import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DepartmentSelectedCallback } from './DepartmentSelectedCallBack';

export interface IProviderWebPartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  siteUrl: string;
  onDepartmentSelected?: DepartmentSelectedCallback;
}
