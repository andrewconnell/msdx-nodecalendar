# Microsoft DX Article - Node.js Authentication with Azure AD for Graph REST API

This sample demonstrates creating a Node.js website using [Express](http://expressjs.com/) & the server-side [handlebars](https://github.com/donpark/hbs) view engine with [Azure AD](https://azure.microsoft.com/en-us/services/active-directory/) to create a simple authentication solution. The application will then demonstrate showing information from a Office 365 calendar using the Microsoft Graph REST API

## Setup

The project is written in TypeScript and leverages the [Typings](https://github.com/typings/typings) utility for type definitions. It is expected that these two things are installed globally in your environment. If not, install them:

  ```shell
  $ npm install -g tsc typings
  ```

Next obtain all packages used by the project:

  ```shell
  $ npm install
  ```

Finally, build everything. An NPM script is provided to make this easy. It will first download all type definitions and then transpile the TypeScript to JavaScript.

  ```shell
  $ npm run build
  ```

> Node: You can also run the application in debug mode. This assumes you have [node-inspector](https://github.com/node-inspector/node-inspector) installed globally.

Next, create an Azure AD application and update the following properties in the `src/server/config.json` file:

- `aad-tenant-id`: The unique ID (GUID) of your Azure AD tenant.
- `aad-client-id`: The unique ID (GUID) of the Azure AD application.
- `aad-client-secret`: The Azure AD application's secret key.
- `msgraph-rest-api-resourceid`: The resource ID for the Microsoft Graph REST API.
- `msgraph-rest-api-endpoint`: The REST API endpoint Microsoft Graph REST API

## Run

To run the application, do the following:

  ```shell
  $ npm start
  ```

> Note: running the application will actually rebuild the project.

The console will display a message for what URLs you can access the web application. Navigate to one of the two URLs. You will see a homepage with two links: **home** & **login**. Clicking **login** will take you to the Azure AD login page. Once logged in it will redirect you back to this application's homepage but it will display your credentials.

Note also that there is a folder on the file system named `session`. Within that folder you will find a JSON file. The session and cookie information for the application is saved to the file system for you to easily examine. No, this isn't the most secure option, but it is useful for learning & demonstration purposes. 