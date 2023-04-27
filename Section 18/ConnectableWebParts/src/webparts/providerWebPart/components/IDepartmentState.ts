import { IDepartment } from './IDepartment';

export interface IProviderWebPartState {
  status: string;
  DepartmentListItems: IDepartment[];
  DepartmentListItem: IDepartment;
}
