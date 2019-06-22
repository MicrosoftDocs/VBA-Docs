---
title: VisWebPageSettings.InitSettings method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.InitSettings
ms.assetid: 11f4d234-53ab-d49c-9c1c-3c8c6ff3f9eb
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.InitSettings method

Loads the webpage settings that were saved in the registry in an earlier instance of Microsoft Visio.


## Syntax

_expression_.**InitSettings**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Nothing**


## Remarks

You can use the **InitSettings** method to reinitialize the webpage settings to those values stored in the registry.


## Example

The following example shows how to use the **InitSettings** method to reinitialize the webpage settings to those that were saved in an earlier instance of Visio.

Before running this example, replace `path\filename` with a valid path and file name for the webpage project file.

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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]