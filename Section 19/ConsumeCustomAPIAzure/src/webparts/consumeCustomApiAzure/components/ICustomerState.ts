
import { ICustomer } from "./ICustomer";

export interface ICustomerState {
    status: string;
    CustomerListItems: ICustomer[];
    CustomerListItem: ICustomer;
  }