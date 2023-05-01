import { ISoftwareListItem } from './ISoftwareListItem';

export interface ICrudReactWebPartState {
  status: string;
  SoftwareListItems: ISoftwareListItem[];
  SoftwareListItem: ISoftwareListItem;
}
