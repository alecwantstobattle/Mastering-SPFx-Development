import {
  Context,
  HttpMethod,
  HttpRequest,
  HttpResponse,
  HttpStatusCode,
} from 'azure-functions-ts-essentials';

const customersDb = require('./customers.json');

const getCustomerById = (customerid: any) => {
  const customer = customersDb.find((customer) => {
    return customer.id === customerid;
  });

  return {
    status: HttpStatusCode.OK,
    body: customer,
  };
};

const getAllCustomers = () => {
  return {
    status: HttpStatusCode.OK,
    body: customersDb,
  };
};

const addNewCustomer = (newCustomer) => {
  const newCustomers = customersDb;
  newCustomers.push(newCustomer);
  return {
    status: HttpStatusCode.Created,
    body: newCustomers,
  };
};

export async function run(context: Context, req: HttpRequest): Promise<any> {
  let response: any;
  const customerid = req.params ? req.params.customerid : undefined;

  switch (req.method) {
    case 'GET':
      response = customerid ? getCustomerById(customerid) : getAllCustomers();
      break;
    case 'POST':
      response = addNewCustomer(req.body);
      break;

    default:
      response = {
        status: HttpStatusCode.BadRequest,
        body: {
          error: {
            type: 'not_supported',
            message: `This Method ${req.method} is not supported.`,
          },
        },
      };
  }

  response.headers = {
    'Content-Type': 'application/json',
  };

  context.res = response;
  Promise.resolve();
}
