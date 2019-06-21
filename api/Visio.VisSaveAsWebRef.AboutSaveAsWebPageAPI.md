---
title: About the Save as Web Page API
ms.prod: visio
ms.assetid: 82d863e2-88a3-527b-4c2e-4c9b43aa3df6
ms.date: 06/20/2019
localization_priority: Normal
---


# About the Save as Web Page API

The Save as Web Page feature, which was introduced in Visio 2002, provides users with a simple way of publishing Visio documents on the web.

The Save as Web Page API gives you programmatic access to this feature, enabling you to save a Visio drawing as a webpage without exposing the user to the **Save as Web Page** dialog boxes in the user interface.

Using this API, you can do the following:

- Save a document as HTML for publishing to the web.    
- Generate the supporting files that are needed to publish your document to the web.   
- View a shape's custom properties in the browser.   
- Display search and navigation controls in the browser.   
- Navigate a multiple-page document in the browser.   
- Display the **Pan and Zoom** control.   
- View all the hyperlinks associated with a shape and navigate to a selected hyperlink target.  
- Assign a webpage a style sheet with color scheme styles that match the color schemes available in Visio.
    

Following are two ways to use the Save as Web Page API.

## Use the API from any development environment that supports Automation 

Using the Save as Web Page object model, you can write code in a document's Visual Basic project, a VSL (a C++ add-on that runs in the Visio address space), or a COM add-in (created with Visual Basic, C++, or C#), and save a Visio drawing as a webpage without any user intervention. 

To control the Save as Web Page feature from an executable that is running in its own process (outside the Visio address space), you can either use the command-line interface, as described next, or you can get a **VisSaveAsWeb** object by using the **SaveAsWebObject** property of the Visio **Application** object. 

For an example of using the Save as Web Page object model in Visual Basic, see [Using the Save as Web Page object model from Visual Basic: an example](Visio.VisSaveAsWebRef.UseSaveAsWebObjectModelVBExample.md). 
    
## Use the API command-line interface 

You can use the command-line interface to call the SaveAsWeb add-on from an executable or from code that is running in the Visio process. Additionally, you can create formulas in the ShapeSheet window that launch the SaveAsWeb add-on without writing any code. For information about using the command-line interface to run the SaveAsWeb add-on, see [Running Save as Web Page from the command line](Visio.VisSaveAsWebRef.RunSaveAsWebPageFromCommandLine.md).
    
Whether you run the Save as Web Page feature from the user interface, from code, or from the command-line interface, the Save as Web Page feature stores selected customized webpage settings in the registry. This data is persisted between instances of Visio and enables users to manage default values for their own projects.

For information about the data that is stored in the registry, see [Persisting Save as Web Page settings](Visio.VisSaveAsWebRef.PersistSaveAsWebPageSettings.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]