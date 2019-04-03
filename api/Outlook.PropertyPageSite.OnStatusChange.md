---
title: PropertyPageSite.OnStatusChange method (Outlook)
keywords: vbaol11.chm389
f1_keywords:
- vbaol11.chm389
ms.prod: outlook
api_name:
- Outlook.PropertyPageSite.OnStatusChange
ms.assetid: d314f8fc-33f5-0a6f-22c0-e26548e21a4f
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyPageSite.OnStatusChange method (Outlook)

Notifies Microsoft Outlook that a custom property page has changed.


## Syntax

_expression_. `OnStatusChange`

_expression_ A variable that represents a [PropertyPageSite](Outlook.PropertyPageSite.md) object.


## Example

This Microsoft Visual Basic for Applications (VBA) example shows how to call the  **[OnStatusChange](Outlook.PropertyPageSite.OnStatusChange.md)** method to notify Outlook that the user has changed a value on a custom property page.


```vb
Private Sub Option1_Click() 
 
 Dim myPPSite As Outlook.PropertyPageSite 
 
 Set myPPSite = Parent 
 
 If Not TypeName(myPPSite) = "Nothing" Then 
 
 globNewUserType = globAdministrator 
 
 If globUserType <> globNewUserType Then 
 
 globDirty = True 
 
 myPPSite.OnStatusChange 
 
 End If 
 
 Else 
 
 If TypeName(myPPSite) = "Nothing" Then 
 
 MsgBox "The Property Page returned an empty result." 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[PropertyPageSite Object](Outlook.PropertyPageSite.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]