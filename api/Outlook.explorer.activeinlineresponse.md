---
title: Explorer.ActiveInlineResponse property (Outlook)
keywords: vbaol11.chm3595
f1_keywords:
- vbaol11.chm3595
ms.assetid: fc38314d-7cff-44f4-9151-6129f918a721
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# Explorer.ActiveInlineResponse property (Outlook)
Returns an item object representing the active inline response item in the explorer reading pane. Read-only.

## Syntax

_expression_. `ActiveInlineResponse`

_expression_ A variable that represents an '[Explorer](Outlook.Explorer.md)' object.


## Remarks

You can use the same properties and methods of the [MailItem](Outlook.MailItem.md) object on this item, except for the following:


- [MailItem.Actions](Outlook.MailItem.Actions.md) property
    
- [MailItem.Close](Outlook.MailItem.Close(method).md) method
    
- [MailItem.Copy](Outlook.MailItem.Copy.md) method
    
- [MailItem.Delete](Outlook.MailItem.Delete.md) method
    
- [MailItem.Forward](Outlook.MailItem.Forward(method).md) method
    
- [MailItem.Move](Outlook.MailItem.Move.md) method
    
- [MailItem.Reply](Outlook.MailItem.Reply(method).md) method
    
- [MailItem.ReplyAll](Outlook.MailItem.ReplyAll(method).md) method
    
- [MailItem.Send](Outlook.MailItem.Send(method).md) method
    
This property returns  **Null** (**Nothing** in Visual Basic) if no inline response is visible in the Reading Pane.


## See also


[Explorer Object](Outlook.Explorer.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]