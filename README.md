## sp-extension-sendto

This is where you include your WebPart documentation.

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

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO

`?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"a0784cae-d0ef-4a7e-9cf0-33e4d31bc759":{"location":"ClientSideExtension.ListViewCommandSet","properties":{"listFieldTitle":"<list-field-title>", "listFieldsEmail": ["<list-field-email-one>","<list-field-email-two>"]}}}`


gulp bundle - TODO
gulp package-solution - TODO
