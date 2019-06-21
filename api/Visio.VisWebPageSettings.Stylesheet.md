---
title: VisWebPageSettings.Stylesheet property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.Stylesheet
ms.assetid: 9b837460-83a6-71f8-b63f-3f251dedc87c
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.Stylesheet property

Specifies a cascading stylesheet (CSS) provided by Microsoft Visio, or one that you have created, that is applied to the webpage. Read/write.


## Syntax

_expression_.**Stylesheet**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**String**


## Remarks

A stylesheet can be one provided by Visio or one that you create yourself. If you store a stylesheet that you create in the following folder, it appears in the **Style sheet** drop-down list on the **Advanced** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish** > **Advanced**): 

> \ _your\_Visio\_path_\ _your\_language\_ID_\

Visio identifies stylesheets by searching through the folder named for your language ID (for example, 1033) for CSS files.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **Stylesheet** property to assign the Steel stylesheet (supplied by Visio) to the webpage that you are creating.

Before running this macro, replace `path\filename.htm` with a valid target path on your computer and the file name that you want to assign to your webpage. Also, replace `your_Visio_path` and `your_language_ID` with the path to Visio stylesheets on your computer, for example:

> C:\Program Files\Microsoft Office\Visio14\1033...

```vb
Public Sub Stylesheet_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .Stylesheet = "your_Visio_path\your_language_ID\Steel.css" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]