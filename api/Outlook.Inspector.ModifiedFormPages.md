---
title: Inspector.ModifiedFormPages property (Outlook)
keywords: vbaol11.chm2964
f1_keywords:
- vbaol11.chm2964
ms.prod: outlook
api_name:
- Outlook.Inspector.ModifiedFormPages
ms.assetid: ac377d47-846a-1217-592f-7ed190b824ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.ModifiedFormPages property (Outlook)

Returns the  **[Pages](Outlook.Pages.md)** collection that represents all the pages for the item in the inspector. Read-only.


## Syntax

_expression_. `ModifiedFormPages`

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Remarks

The main page and up to five customizable pages can be obtained using the  **[Add](Outlook.Pages.Add.md)** method.


## Example

This Visual Basic for Applications (VBA) displays the count of pages in the  **ModifiedFormPages** collection. To run this example without any errors, display a contact item in the active window.


```vb
Sub CountModifiedFormPages() 
 
 Dim myItem As Outlook.ContactItem 
 
 Dim myPages As Outlook.Pages 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 Set myPages = myItem.GetInspector.ModifiedFormPages 
 
 MsgBox myPages.Count 
 
End Sub
```


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]