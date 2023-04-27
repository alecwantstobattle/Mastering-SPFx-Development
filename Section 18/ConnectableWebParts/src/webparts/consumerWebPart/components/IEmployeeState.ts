import { IEmployee } from './IEmployee';

export interface IConsumerWebPartState {
  status: string;
  EmployeeListItems: IEmployee[];
  EmployeeListItem: IEmployee;
  DeptTitleId: string;
}
