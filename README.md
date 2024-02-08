
## Sample goal

This sample will check out if the current SharePoint user has already read the site disclamer

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## HOW TO RUN

1. Open repository folder on Visual Studio Code
2. Go to "config/serve.json" file and uptate "initialPage" value to your SharePoint URL (for test purposes it's recommended to keep "/_layouts/workbench.aspx" last part of URL)
3. Make sure you have at least one SharePoint list created (it only has to contains the Title column) where WebPart will read and write
4. Open a terminal
5. Run "npm install"
6. Run "gulp trust-dev-cert"
7. Run "gulp serve" command.
8. Once your SharePoint page opens in the browser, you must use "+" icon com add your custom WebPart (search for your WebPart name: e.g "hello")
