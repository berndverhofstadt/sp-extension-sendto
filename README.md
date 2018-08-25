## sp-extension-sendto

This is where you include your Extension documentation.

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Debug URL for testing
Here's a debug querystring for testing this sample:

```
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"a0784cae-d0ef-4a7e-9cf0-33e4d31bc759":{"location":"ClientSideExtension.ListViewCommandSet","properties":{"defaultTo":"info@example.org", "listFieldsEmail": ["<list-field-email-one>","<list-field-email-two>"]}}}
```

Your URL will look similar to the following (replace with your domain and site address):
```
https://yourtenant.sharepoint.com/sites/yoursite/Lists/yourlist/AllItems.aspx?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"a0784cae-d0ef-4a7e-9cf0-33e4d31bc759":{"location":"ClientSideExtension.ListViewCommandSet","properties":{"defaultTo":"info@example.org", "listFieldsEmail": ["<list-field-email-one>","<list-field-email-two>"]}}}
```

## Deploying to your tenant
- In the command line navigate to **path/sp-extension-emailto** and run:
  - `gulp bundle --ship`
  - `gulp package-solution --ship`
- Open the **path/sp-extension-emailto/sharepoint** folder
  - Drag the **sp-extension-emailto.sppkg** onto the **Apps for SharePoint** library of your app catalog
- You'll need to add the Custom Action to your site(s) using one of the methods below

### Adding the custom action to your site

Even if you selected tenant wide deployment for the package, each site will need a Custom Action added to take advantage of the extension.

### Option 2: Use the SPFx Extensions CLI
You can use the [spfx-extensions-cli](https://www.npmjs.com/package/spfx-extensions-cli) to manage your extension custom actions across your sites.

Install the CLI if you haven't already:

`npm install spfx-extensions-cli -g`

Connect to your site (login when prompted):

`spfx-ext --connect "https://yourtenant.sharepoint.com/sites/yoursite"`

Add the extension (be sure to replace the listtitle parameter with the name of your list):

`spfx-ext add "Send Email" ListViewCommandSet list a0784cae-d0ef-4a7e-9cf0-33e4d31bc759 --listtitle "<list-name>" --clientProps  {\"defaultTo\":\"info@example.org\", \"listFieldsEmail\": [\"<email-field-one>\",\"<email-field-two>\"]}"`

## Known issues
- UI Fabric Icons are not currently displaying in SPFx Extensions: 
  - [Issue 1279](https://github.com/SharePoint/sp-dev-docs/issues/1279) - Solution has been found, but fix has not yet been implemented
