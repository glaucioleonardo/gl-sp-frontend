# gl-sp-frontend #
[![version](https://img.shields.io/npm/v/gl-sp-frontend)](https://www.npmjs.com/package/gl-sp-frontend) 
[![downloads](https://img.shields.io/npm/types/gl-sp-frontend)](https://www.npmjs.com/package/gl-sp-frontend) 
[![downloads](https://img.shields.io/npm/dw/gl-sp-frontend)](https://www.npmjs.com/package/gl-sp-frontend)
 
### What is this repository for? ###
This package (**g**laucio**l**eonardo-**s**hare**p**oint-**frontend**) is a common code for SharePoint 2013/2016 using among web front-end development such as pure 
JavaScript, ES6+ and TypeScript.<br>

This package contains the types for using with Typescript.

### How do I get set up? ###
 The easiest way to install this library is via npm using the following commands:
* Latest version `npm install gl-sp-frontend --save`;
* [Available versions](https://www.npmjs.com/package/gl-sp-frontend?activeTab=versions) `npm install gl-sp-frontend@version --save`;
* If you need to support old browsers (tested in IE10+), just install those packages polyfills and import as 
the following sequence:
  * [es6-promise](https://github.com/stefanpenner/es6-promise): `npm install es6-promise --save`
  * [whatwg-fetch](https://github.com/github/fetch): `npm install whatwg-fetch --save`

This is how you should include in your code and voila \o/:

```
import * as promise from 'es6-promise';
import 'whatwg-fetch';

class MyBeautifulClass {
    constructor() {
        promise.polyfill();
    }
}
```

If you are using Angular 2+, just import these packages inside polyfill.js and in the section:<br>
`/** IE10 and IE11 requires the following for external source of SVG when using <use> tag */`<br> 
include `promise.polyfill();` and everything should work fine!

If you are using only browser version:
* For ES5 version importing via <br>
```
<!DOCTYPE html>
 <html lang="en">
 <head>
    <meta charset="UTF-8">
    <title>gl-sp-frontend</title>
    ...
 </head>
 <body>
    ...
    <!-- Include here -->
    <script src="https://github.com/glaucioleonardo/gl-sp-frontend/tree/master/lib/bundle.js"></script>
    <!-- or just download this repo and refer to the bundle.js like this -->
    <script src="js/bundle.js"></script>
 </body>
 </html>
 ```

* In case you are using ES+, just use the `bundle.js` inside `lib/es6`

### Macro features ###

* Core: Setup
* Retrieve list items
* User permissions 

Other features is going to be included  frequently.

### Other packages used ###
I'm really thankful for those packages creators!
* [PnP-JS-Core](https://github.com/SharePoint/PnP-JS-Core)
