---
title: VisWebPageSettings.ThemeName property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.ThemeName
ms.assetid: 9efd26b1-7426-1ff4-0b51-5463a2beb822
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.ThemeName property

Assigns a webpage theme to the page that you are creating. Read/write.


## Syntax

_expression_.**ThemeName**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**String**


## Remarks

You can use themes that are provided by Microsoft Visio or themes that you create yourself. If you want to create your own theme, do the following: 

1. Create an HTM file that contains the following term in an HTML tag: `##VIS_SAW_FILE##`. Visio recognizes HTM files that contain this tag as theme files.
    
2. Store the file in the following folder: \ _your_Visio_path_\ _your_language_ID_\
    
Your theme file will then appear in the **Host in Web page** drop-down list in the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish** > **Advanced**).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **ThemeName** property to assign the Basic theme (supplied by Visio) to the webpage that you are creating.

Before running this macro, replace `path\filename.htm` with a valid target path on your computer and the file name that you want to assign to your webpage. Also, replace `your_Visio_path` and `your_language_ID` with the path to Microsoft Visio on your computer (for example, C:\Program Files\Microsoft Office\Visio14\1033\...).

```vb
Public Sub ThemeName_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsoWebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .ThemeName = "your_Visio_path\your_language_ID\Basic.htm" 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]