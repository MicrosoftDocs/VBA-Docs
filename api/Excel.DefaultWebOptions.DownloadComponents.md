---
title: DefaultWebOptions.DownloadComponents property (Excel)
keywords: vbaxl10.chm660080
f1_keywords:
- vbaxl10.chm660080
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.DownloadComponents
ms.assetid: 8522baf4-77da-4e0b-30b1-604a2a4493d0
ms.date: 04/25/2019
localization_priority: Normal
---


# DefaultWebOptions.DownloadComponents property (Excel)

**True** if the necessary Microsoft Office Web components are downloaded when you view the saved document in a web browser, but only if the components are not already installed. **False** if the components are not downloaded. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**DownloadComponents**

_expression_ A variable that represents a **[DefaultWebOptions](Excel.DefaultWebOptions.md)** object.


## Remarks

You can set the **[LocationOfComponents](Excel.DefaultWebOptions.LocationOfComponents.md)** property to a central URL (on the intranet or web) or path (local or network) to a location from which authorized users can download components when viewing your saved document. The path must be valid and must point to a location that contains the necessary components, and the user must have a valid Microsoft Office license.

Office Web components add interactivity to documents that you save as webpages. If you view a webpage in a browser on a computer that does not have the components installed, the interactive portions of the page will be static.


## Example

This example allows the Office Web components to be downloaded with the specified webpage, if they are not already installed.

```vb
Application.DefaultWebOptions.DownloadComponents = True 
Application.DefaultWebOptions.LocationOfComponents = _ 
 Application.Path & Application.PathSeparator & "foo"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]