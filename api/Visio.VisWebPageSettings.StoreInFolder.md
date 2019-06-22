---
title: VisWebPageSettings.StoreInFolder property (Visio Save As Web)
ms.prod: visio
api_name:
- Visio.VisWebPageSettings.StoreInFolder
ms.assetid: ed0cf76a-a68d-cfa7-538c-91df5234a0d0
ms.date: 06/21/2019
localization_priority: Normal
---


# VisWebPageSettings.StoreInFolder property

Determines whether supporting files for the webpage to be created are placed into a subfolder that has the same name as the root HTML file. Read/write.


## Syntax

_expression_.**StoreInFolder**

_expression_ An expression that returns a **[VisWebPageSettings](Visio.VisWebPageSettings.md)** object.


## Return value

**Long**


## Remarks

Set **StoreInFolder** to a non-zero value (**True**) to place supporting webpage files in a subfolder that has the same name as the root HTML file; otherwise, set it to zero (**False**). 

If you set the **StoreInFolder** property to **True** (non-zero), Microsoft Visio places the supporting files in a subfolder prefixed with the same name as the .htm file. If either the .htm file or the subfolder is moved or deleted, its corresponding subfolder or .htm file is also moved or deleted.

If you set the **StoreInFolder** property to **False** (0), Visio places all supporting files in the same folder as the .htm file.

Setting the **StoreInFolder** property to **True** is the equivalent of selecting the **Organize supporting files in a folder** check box on the **General** tab of the **Save As Web Page** dialog box (**BackstageButton** tab > **Save As** > **Save as type** list > **Web Page (\*.htm;\*.html)** > **Publish**).


## Example

The following macro shows how to set the **StoreInFolder** property so that a subfolder that contains all of a webpage's supporting files and has the same name as the .htm file is created.

Before running this macro, replace `path\filename.htm` with a valid target path on your computer and the filename that you want to assign to your webpage.

```vb
Public Sub StoreInFolder_Example() 
 Dim vsoSaveAsWeb As VisSaveAsWeb 
 Dim vsowebSettings As VisWebPageSettings 
 
 Set vsoSaveAsWeb = Visio.Application.SaveAsWebObject 
 Set vsoWebSettings = vsoSaveAsWeb.WebPageSettings 
 
 With vsoWebSettings 
 .StoreInFolder = True 
 .TargetPath = "path\filename.htm" 
 End With 
 
 vsoSaveAsWeb.CreatePages 
 End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]