
# react-native-office-document-tool

## Getting started

`$ npm install react-native-office-document-tool --save`

### Mostly automatic installation

`$ react-native link react-native-office-document-tool`

### Manual installation


#### iOS

1. In XCode, in the project navigator, right click `Libraries` ➜ `Add Files to [your project's name]`
2. Go to `node_modules` ➜ `react-native-office-document-tool` and add `RNOfficeDocumentTool.xcodeproj`
3. In XCode, in the project navigator, select your project. Add `libRNOfficeDocumentTool.a` to your project's `Build Phases` ➜ `Link Binary With Libraries`
4. Run your project (`Cmd+R`)<

#### Android

1. Open up `android/app/src/main/java/[...]/MainActivity.java`
  - Add `import com.reactlibrary.RNOfficeDocumentToolPackage;` to the imports at the top of the file
  - Add `new RNOfficeDocumentToolPackage()` to the list returned by the `getPackages()` method
2. Append the following lines to `android/settings.gradle`:
  	```
  	include ':react-native-office-document-tool'
  	project(':react-native-office-document-tool').projectDir = new File(rootProject.projectDir, 	'../node_modules/react-native-office-document-tool/android')
  	```
3. Insert the following lines inside the dependencies block in `android/app/build.gradle`:
  	```
      compile project(':react-native-office-document-tool')
  	```

#### Windows
[Read it! :D](https://github.com/ReactWindows/react-native)

1. In Visual Studio add the `RNOfficeDocumentTool.sln` in `node_modules/react-native-office-document-tool/windows/RNOfficeDocumentTool.sln` folder to their solution, reference from their app.
2. Open up your `MainPage.cs` app
  - Add `using Office.Document.Tool.RNOfficeDocumentTool;` to the usings at the top of the file
  - Add `new RNOfficeDocumentToolPackage()` to the `List<IReactPackage>` returned by the `Packages` method


## Usage
```javascript
import RNOfficeDocumentTool from 'react-native-office-document-tool';

// TODO: What to do with the module?
RNOfficeDocumentTool;
```
  