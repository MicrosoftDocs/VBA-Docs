---
title: VisWebPageSettings.SaveSettings Method (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.SaveSettings
ms.assetid: c3b7ba3c-23a0-285f-c668-d220e9d99833
ms.date: 06/08/2017
localization_priority: Normal
---


# VisWebPageSettings.SaveSettings Method (Visio Save As Web)

Saves the current Web page settings to the registry.


## Syntax

 _expression_. **SaveSettings**

 _expression_An expression that returns a  ** [VisWebPageSettings](./overview/Visio.md)** object.


## Return value

 **Nothing**


## Remarks

By default, when some Web page settings are explicitly set to something other than the default value, they are saved to the registry when a Save as Web Page project's files are exported to the target path. The  **SaveSettings** method causes these settings to be written to the registry when the method is called rather than waiting until the files are exported.

For more information about which settings are persisted to the registry, see  [Persisting Save as Web Page Settings](Visio.persisting.save.as.web.md).


## Example

The following example shows how to use the  **SaveSettings** method to immediately change the default value for the **[PriFormat](Visio.PriFormat.md)** property.

Before running this example, replace  _path\filename_ with a valid path and file name for the Web page project files.




```vb
Public Sub SaveSettings_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 'Set PriFormat to a non-default value. 
 .PriFormat = "JPG" 
 .SaveSettings 
 .TargetPath = "path\filename" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]