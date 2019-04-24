---
title: WebOptions.DownloadComponents property (Excel)
keywords: vbaxl10.chm662076
f1_keywords:
- vbaxl10.chm662076
ms.prod: excel
api_name:
- Excel.WebOptions.DownloadComponents
ms.assetid: d9f103f8-e41e-ee8b-0e02-8cda514f04c9
ms.date: 06/08/2017
localization_priority: Normal
---


# WebOptions.DownloadComponents property (Excel)

 **True** if the necessary Microsoft Office Web components are downloaded when you view the saved document in a Web browser, but only if the components are not already installed. **False** if the components are not downloaded. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**DownloadComponents**

_expression_ A variable that represents a **[WebOptions](Excel.WebOptions.md)** object.


## Remarks

You can set the  **[LocationOfComponents](Excel.WebOptions.LocationOfComponents.md)** property to a central URL (on the intranet or Web) or path (local or network) to a location from which authorized users can download components when viewing your saved document. The path must be valid and must point to a location that contains the necessary components, and the user must have a valid Microsoft Office license.

Office Web components add interactivity to documents that you save as web pages. If you view a web page in a browser on a computer that does not have the components installed, the interactive portions of the page will be static.


## Example

This example allows the Office Web components to be downloaded with the specified web page, if they are not already installed.


```vb
Application.DefaultWebOptions.DownloadComponents = True 
Application.DefaultWebOptions.LocationOfComponents = _ 
 Application.Path & Application.PathSeparator & "foo"
```


## See also


[WebOptions Object](Excel.WebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]