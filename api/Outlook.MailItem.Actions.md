---
title: MailItem.Actions property (Outlook)
keywords: vbaol11.chm1294
f1_keywords:
- vbaol11.chm1294
ms.prod: outlook
api_name:
- Outlook.MailItem.Actions
ms.assetid: 1b7bb1c0-334f-826a-fd6b-8fc3f2fe5d64
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Actions property (Outlook)

Returns an  **[Actions](Outlook.Actions.md)** collection that represents all the available actions for the item. Read-only.


## Syntax

_expression_. `Actions`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Example

This Visual Basic for Applications (VBA) example creates a new mail item and uses the  **[Actions.Add](Outlook.Actions.Add.md)** method to add an **[Action](Outlook.Action.md)** to it. Then it sends the mail item to the current user. The mail item received will have the **Agree** action in addition to the standard actions such as **Reply** and **Reply All**.


```vb
Sub AddAction() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myAction As Outlook.Action 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myAction = myItem.Actions.Add 
 
 myAction.Name = "Agree" 
 
 myItem.To = Application.GetNamespace("MAPI").CurrentUser 
 
 myItem.Send 
 
End Sub
```

The following Visual Basic for Applications example creates a new mail item and uses the  **Actions.Add** method to add an **Action** called **Link Original** to it. Executing this action will insert a link to the original mail item.




```vb
Sub AddAction2() 
 
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


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]