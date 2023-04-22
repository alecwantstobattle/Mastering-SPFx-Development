import { IUser } from './IUser';

export interface IShowAllUsersOfficeFabricState {
  users: Array<IUser>;
  searchFor: string;
}
