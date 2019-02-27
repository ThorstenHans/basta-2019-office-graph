# Office Add-in with Microsoft Graph

In order to get that Office Add-in up and running, you've to setup an Azure AD Application Registration. 

Provide the following settings:

* Supported Account Types: Accounts in any organizational directory
* Redirect URI: https://localhost:4200 or the https URI of your choice
* API Permissions `User.Read` and `Files.Read` from *Microsoft Graph*
* Enable `Authentication` -> `Implicit Grant`

Grab the `Client ID` from the Azure AD App Registration you've just created and set it in
`environment.aad.clientId` (see `environment.ts` and `environment.prod.ts`)

Either provide your own certificate or trust the developer cert for localhost.

Start the web app using `npm start`

Sideload the Office Add-in either using `office-toolbox` or manually as described on dev.office.com

