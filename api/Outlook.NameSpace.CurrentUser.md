---
title: NameSpace.CurrentUser property (Outlook)
keywords: vbaol11.chm756
f1_keywords:
- vbaol11.chm756
ms.prod: outlook
api_name:
- Outlook.NameSpace.CurrentUser
ms.assetid: d6884fcf-c1de-23f4-8d91-02c8f9fd5253
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.CurrentUser property (Outlook)

Returns the display name of the currently logged-on user as a **[Recipient](Outlook.Recipient.md)** object. Read-only.


## Syntax

_expression_. `CurrentUser`

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Example

This Visual Basic for Applications (VBA) example uses the  **CurrentUser** property to obtain the name of the currently logged-on user and then displays a message box containing the name.


```vb
Sub DisplayCurrentUser() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 MsgBox myNameSpace.CurrentUser 
 
End Sub
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
