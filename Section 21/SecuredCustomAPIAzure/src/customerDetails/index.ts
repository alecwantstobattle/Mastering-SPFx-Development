import {
  Context,
  HttpMethod,
  HttpRequest,
  HttpResponse,
  HttpStatusCode,
} from 'azure-functions-ts-essentials';

import jwt = require('jsonwebtoken');

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

const decodedValidToken = (accessToken: string) => {
  const key: string =
    '-----BEGIN CERTIFICATE-----\nMIIDBTCCAe2gAwIBAgIQN33ROaIJ6bJBWDCxtmJEbjANBgkqhkiG9w0BAQsFADAtMSswKQYDVQQDEyJhY2NvdW50cy5hY2Nlc3Njb250cm9sLndpbmRvd3MubmV0MB4XDTIwMTIyMTIwNTAxN1oXDTI1MTIyMDIwNTAxN1owLTErMCkGA1UEAxMiYWNjb3VudHMuYWNjZXNzY29udHJvbC53aW5kb3dzLm5ldDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKGiy0/YZHEo9rRn2bI27u189Sq7NKhInFz5hLCSjgUB2rmf5ETNR3RJIDiW1M51LKROsTrjkl45cxK6gcVwLuEgr3L1TgmBtr/Rt/riKyxeXbLQ9LGBwaNVaJrSscxfdFbJa5J+qzUIFBiFoL7kE8ZtbkZJWBTxHEyEcNC52JJ8ydOhgvZYykete8AAVa2TZAbg4ECo9+6nMsaGsSBncRHJlRWVycq8Q4HV4faMEZmZ+iyCZRo2fZufXpn7sJwZ7CEBuw4qycHvUl6y153sUUFqsswnZGGjqpKSq7I7sVI9vjB199RarHaSSbDgL2FxjmASiUY4RqxnTjVa2XVHUwUCAwEAAaMhMB8wHQYDVR0OBBYEFI5mN5ftHloEDVNoIa8sQs7kJAeTMA0GCSqGSIb3DQEBCwUAA4IBAQBnaGnojxNgnV4+TCPZ9br4ox1nRn9tzY8b5pwKTW2McJTe0yEvrHyaItK8KbmeKJOBvASf+QwHkp+F2BAXzRiTl4Z+gNFQULPzsQWpmKlz6fIWhc7ksgpTkMK6AaTbwWYTfmpKnQw/KJm/6rboLDWYyKFpQcStu67RZ+aRvQz68Ev2ga5JsXlcOJ3gP/lE5WC1S0rjfabzdMOGP8qZQhXk4wBOgtFBaisDnbjV5pcIrjRPlhoCxvKgC/290nZ9/DLBH3TbHk8xwHXeBAnAjyAqOZij92uksAv7ZLq4MODcnQshVINXwsYshG1pQqOLwMertNaY5WtrubMRku44Dw7R\n-----END CERTIFICATE-----';

  return jwt.verify(accessToken, key);
};

export async function run(context: Context, req: HttpRequest): Promise<any> {
  let response: any;
  const customerid = req.params ? req.params.customerid : undefined;

  let blValidRequest: boolean = false;
  let blCustomerReadScope: boolean = false;
  let blCustomerWriteScope: boolean = false;
  let isUser: boolean = false;
  const authorizationHeader: string = req.headers.authorization;

  try {
    const decodedToken = decodedValidToken(
      authorizationHeader.replace('Bearer ', '')
    ) as any;
    console.log('decoded token is : ' + decodedToken);

    const allScopes: string = decodedToken.scp as string;

    blCustomerReadScope = allScopes.indexOf('Customer.Read') >= 0;
    blCustomerWriteScope = allScopes.indexOf('Customer.Write') >= 0;

    isUser = decodedToken.upn.indexOf('sample@sample.com') !== -1;

    blValidRequest = true;
  } catch (err) {
    blValidRequest = false;
    switch (err.name) {
      case 'NotBeforeError':
        response = {
          status: HttpStatusCode.Unauthorized,
          body: {
            message: `${err.message} : ${err.date}`,
          },
        };
        break;
      case 'TokenExpiredError':
        response = {
          status: HttpStatusCode.Unauthorized,
          body: {
            message: `${err.message} : ${err.expiredAt}`,
          },
        };
        break;
      case 'JsonWebTokenError':
        response = {
          status: HttpStatusCode.Unauthorized,
          body: {
            message: `${err.message}`,
          },
        };
        break;
      default:
        response = {
          status: HttpStatusCode.Unauthorized,
          body: {
            message: `Error while decoding & validating json web token: ${err.message}`,
          },
        };
        break;
    }
  }

  if (blValidRequest) {
    switch (req.method) {
      case 'GET':
        if (blCustomerReadScope) {
          response = customerid
            ? getCustomerById(customerid)
            : getAllCustomers();
        } else {
          response = {
            status: HttpStatusCode.Unauthorized,
            body: {
              message:
                'You must have scope of Customer.Read to get customers details.',
            },
          };
        }
        break;
      case 'POST':
        if (blCustomerWriteScope) {
          response = addNewCustomer(req.body);
        } else {
          response = {
            status: HttpStatusCode.Unauthorized,
            body: {
              message: 'You must have Customer.Write scope to add customers.',
            },
          };
        }
        break;

      default:
        response = {
          status: HttpStatusCode.BadRequest,
          body: {
            error: {
              type: 'not_supported',
              message: `Currently this method ${req.method} is not supported.`,
            },
          },
        };
    }
  }

  response.headers = {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Credentials': 'true',
  };

  context.res = response;
  Promise.resolve();
}
