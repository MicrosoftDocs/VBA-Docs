---
title: VisWebPageSettings.EndPage property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.EndPage
ms.assetid: 4b7ebf2d-b814-8588-b25e-7c54fd0affda
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.EndPage property

Specifies the page number of the last page in the range when you save a range of pages as a webpage. Read/write.


## Syntax

_expression_.**EndPage**

 _expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

The start page number is specified in the **[StartPage](Visio.VisWebPageSettings.StartPage.md)** property.

The **EndPage** property value corresponds to the value in the **to** box on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **EndPage** property to save a range of pages in a drawing (in this case, from page 2 to page 3) as a webpage instead of the complete drawing.

This macro assumes that the current Visio drawing contains at least three pages.

Before running this macro, replace _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your webpage.

```vb
Public Sub EndPage_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .StartPage = 2 
 .EndPage = 3 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]