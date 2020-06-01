# gl-sp-frontend #
[![version](https://img.shields.io/badge/version-1.4.3-green.svg)](https://www.npmjs.com/package/gl-w-frontend)

### What is this repository for? ###
This package is a common code for SharePoint 2013/2016 using among web front-end development such as pure 
JavaScript, ES6+ and TypeScript.<br>

This package contains the types for using with Typescript.

### How do I get set up? ###
 The easiest way to install this library is via npm using the following commands:
* Latest version `npm install gl-sp-frontend --save`;
* [Available versions](https://www.npmjs.com/package/gl-sp-frontend?activeTab=versions) `npm install gl-sp-frontend@version --save`;


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
    <script src="https://github.com/glaucioleonardo/gl-sp-frontend/tree/master/lib/index.js"></script>
    <!-- or just download this repo and refer to the index.js like this -->
    <script src="js/index.js"></script>
 </body>
 </html>
 ```

* In case you are using ES+, just use the `index.js` inside `lib/esm`

### Macro features ###

* Core: Setup
* Retrieve list items 

Other features is going to be included  frequently.

### Other packages used ###
I'm really thankful for those packages creators!
* [PnP-JS-Core](https://github.com/SharePoint/PnP-JS-Core)
