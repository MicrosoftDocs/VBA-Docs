---
title: VisWebPageSettings.InitSettings Method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.InitSettings
ms.assetid: 11f4d234-53ab-d49c-9c1c-3c8c6ff3f9eb
ms.date: 06/08/2017
localization_priority: Normal
---


# VisWebPageSettings.InitSettings Method (Visio Save As Web)

Loads the Web page settings that were saved in the registry in an earlier instance of Microsoft Visio.


## Syntax

 _expression_. **InitSettings**

 _expression_An expression that returns a  ** [VisWebPageSettings](./overview/Visio.md)** object.


## Return value

 **Nothing**


## Remarks

You can use the  **InitSettings** method to reinitialize the Web page settings to those values stored in the registry.


## Example

The following example shows how to use the  **InitSettings** method to reinitialize the Web page settings to those that were saved in an earlier instance of Visio.

Before running this example, replace  _path\filename_ with a valid path and file name for the Web page project file.




```vb
Public Sub InitSettings_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .InitSettings 
 .TargetPath = "path\filename" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```


