export interface ICustomer {
    id: string;
    joined: string;
    name: string;
    city: string;
    orderTotal: number;    
    orders?: any[];    
  }
  