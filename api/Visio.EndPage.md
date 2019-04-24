---
title: VisWebPageSettings.EndPage Property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.EndPage
ms.assetid: 4b7ebf2d-b814-8588-b25e-7c54fd0affda
ms.date: 06/08/2017
localization_priority: Normal
---


# VisWebPageSettings.EndPage Property (Visio Save As Web)

Specifies the page number of the last page in the range when you save a range of pages as a webpage. Read/write.


## Syntax

_expression_.**EndPage**

 _expression_ An expression that returns a  **[VisWebPageSettings](visio.viswebpagesettings.object.visio.save.md)** object.


## Return value

 **Long**


## Remarks

The start page number is specified in the  **[StartPage](Visio.StartPage.md)** property.

The  **EndPage** property value corresponds to the value in the **to** box on the **General** tab of the **Save As Web Page** dialog box (click the **BackstageButton** tab, click **Save As**, in the  **Save as type** list, select **Web Page (*.htm;*.html)**, and then click  **Publish**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **EndPage** property to save a range of pages in a drawing (in this case, from page 2 to page 3) as a webpage instead of the complete drawing.

This macro assumes that the current Visio drawing contains at least three pages.

Before running this macro, replace  _path\filename.htm_ with a valid target path on your computer and the filename that you want to assign to your Web page.




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