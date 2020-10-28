## universal-search

### Setup

Requires Node v10.x

```bash
git clone the repo
npm i
npm i -g gulp
```

### Scripts

- `start` - alias for `gulp serve` to preview the web part
- `deploy` - build, package, and zip the web part & teams manifest for deployment or publish
- `deploy:debug` - build and package the web part for debugging. Use `start` to serve.
- `teams` - zip the teams manifest for deployment
- `lint` - typescript check
- `prettier` - run prettier across files

### Dev Notes

- [SPFx Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Publish checklist](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-checklist)


## Deployment Instructions

### Get Files

Either get the files from [the Now site](https://now.regroove.ca/collections/products/products/super-search) or build the solution by running `npm run deploy`.

### Adding to the App Catalog

1. Go to the SharePoint app catalog, or create one if necessary (see [documentation](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog)).
2. Upload or drag the `.sppkg` solution file to the app catalog.
3. With _Make this solution available to all sites in the organization_ selected, click _Deploy_.

### Approving Permissions
1. Go to the [SharePoint admin centre](https://admin.microsoft.com/sharepoint?page=home&modern=true).
2. Under the _Advanced_ settings on the left side, choose _API Access_.
3. Approve both pending permission requests from _universal-search_. 

### Adding to Teams

1. Navigate to the Apps section of Teams from the left toolbar.
2. Add the bottom, choose _Upload a custom app_ and choose the `teams.zip` manifest file.
3. Add it to Teams to enable its use as a Teams personal app or channel tab.

### Updating

The steps to update in the app catalog are the same, which will update the web part for all platforms. To update the Teams manifest if needed, choose the existing app, and update it from the ellipsis context menu.