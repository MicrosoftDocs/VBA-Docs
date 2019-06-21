---
title: VisWebPageSettings.PriFormat property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.PriFormat
ms.assetid: 84c7c085-0f12-f25d-bf17-646cc8b7cd97
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.PriFormat property

Specifies the primary output format for the webpage. Read/write.


## Syntax

_expression_.**PriFormat**

 _expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**String**


## Remarks

If you select a primary output format that is not supported by all browsers, you should also select a secondary output format for older browsers. To do this, see the **[SecFormat](Visio.VisWebPageSettings.SecFormat.md)** property.

For information about which browsers are compatible with selected formats, see the **[AltFormat](Visio.VisWebPageSettings.AltFormat.md)** property.

Possible values for the **PriFormat** property are as follows:

- XAML (Extensible Application Markup Language), the default    
- SVG (Scalable Vector Graphics)    
- JPG (JPEG File Interchange Format)    
- GIF (Graphics Interchange Format)    
- PNG (Portable Network Graphics)    
- VML (Vector Markup Language)
    
This value corresponds to the value selected in the **Output formats** list on the **Advanced** tab of the **Save as Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish** > **Advanced**).


## Example

The following macro shows how to use the **PriFormat** property to set the primary output format for the webpage to JPG.

Before running this macro, replace `path\filename.htm` with a valid target path on your computer and the filename that you want to assign to your webpage.

```vb
Public Sub PriFormat_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .PriFormat = "JPG" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]