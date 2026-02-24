# Setup
* Create a google apps script project
  * Make note of its id
* `yarn install`
* Create *.clasp.json.dev* and/or *.clasp.json.prod*.
  * `{"scriptId":"{your-scriptid-here}"}`
* `yarn run deploy:dev` / `yarn run deploy:dev`
# Usage
* Open a google sheet
* Addons → Apps Script
* Libraries + → search for your project's id
* Add it as "BookEditHelper"
* Go to the code (often `code.gs`).
* Add `BookEditHelper.Setup();` to onOpen function
```typescript
function onOpen() {
  BookEditHelper.Setup();
}
```

# Creds
* Word count function: https://stackoverflow.com/a/78375674