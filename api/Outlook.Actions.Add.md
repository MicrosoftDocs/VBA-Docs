---
title: Actions.Add method (Outlook)
keywords: vbaol11.chm151
f1_keywords:
- vbaol11.chm151
ms.prod: outlook
api_name:
- Outlook.Actions.Add
ms.assetid: aaf539c4-d60a-867f-086b-3cef7632a6f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Actions.Add method (Outlook)

Creates a new action in the  **[Actions](Outlook.Actions.md)** collection.


## Syntax

_expression_.**Add**

_expression_ A variable that represents an [Actions](Outlook.Actions.md) object.


## Return value

An  **[Action](Outlook.Action.md)** object that represents the new action.


## Example

This VBA example creates a new mail message and uses the  **Add** method to add an **[Action](Outlook.Action.md)** to it. To run this example without any errors, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub AddAction() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myAction As Outlook.Action 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myAction = myItem.Actions.Add 
 
 myAction.Name = "Link Original" 
 
 myAction.ShowOn = olMenuAndToolbar 
 
 myAction.ReplyStyle = olLinkOriginalItem 
 
 myItem.To = "Dan Wilson" 
 
 myItem.Send 
 
End Sub
```


## See also


[Actions Object](Outlook.Actions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]