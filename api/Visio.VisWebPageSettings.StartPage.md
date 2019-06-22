---
title: VisWebPageSettings.StartPage property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.StartPage
ms.assetid: 7db581ab-f656-f97a-79b6-17a1fca513e8
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.StartPage property

Specifies the page number of the first page in the range when you save a range of document pages as a webpage. Read/write.


## Syntax

_expression_.**StartPage**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

The end page number is specified in the **[EndPage](Visio.VisWebPageSettings.EndPage.md)** property.

The **StartPage** property value corresponds to the value in the **From** box on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **StartPage** property to save a range of pages in a drawing (in this case, from page 2 to page 3) as a webpage instead of the complete drawing.

This macro assumes that the current Microsoft Visio drawing contains at least three pages.

Before running this macro, replace `path\filename.htm` with a valid target path on your computer and the filename that you want to assign to your webpage.

```vb
Public Sub StartPage_Example() 
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