## hosting-through-local-server-demo

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
gulp bundle - TODO
gulp package-solution - TODO

### Commands
SharePoint Online Management Shell
https://www.microsoft.com/en-us/download/details.aspx?id=35588

Connect-SPOService -Url https://contoso-admin.sharepoint.com

Set-SPOTenantCdnEnabled -CdnType Public

Get-SPOTenantCdnEnabled -CdnType Public
Get-SPOTenantCdnOrigins -CdnType Public
Get-SPOTenantCdnPolicies -CdnType Public


New-SPOPublicCdnOrigin -Url https://kameswarasarma.sharepoint.com/SiteAssets/HelloWorldWPCDN

Get-SPOPublicCdnOrigins | Format-List

https://publiccdn.sharepointonline.com/kameswarasarma.sharepoint.com/159000648534d21b312c47e7495ce8780ada3a81c02b0009f2e44a3b7c53ba72e6d71f7c

Stroage Account Name : spfxstorageaccountit
Container : spfxcontainer

access key: 8iC1FqEuTSWSN6pEWubdaB4wGeTQjLEaXIW8x2qZRPvVduFvw4wrso9UEy73wVHyxlND6WkDWUTllhO5Ev+XGA==

spfxstorageaccountit.blob.core.windows.net



