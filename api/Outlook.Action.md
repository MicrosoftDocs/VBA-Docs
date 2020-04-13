---
title: Action object (Outlook)
keywords: vbaol11.chm9
f1_keywords:
- vbaol11.chm9
ms.prod: outlook
api_name:
- Outlook.Action
ms.assetid: 22bd8d4a-9cf4-bd37-011b-8da3dfadf761
ms.date: 06/08/2017
localization_priority: Normal
---


# Action object (Outlook)

Represents a specialized action (for example, the voting options response) that can be executed on an Outlook item.


## Remarks

The **Action** object is a member of the **[Actions](Outlook.Actions.md)** collection.

Use  **[Actions](Outlook.MailItem.Actions.md)** (_index_), where _index_ is the name of an available action, to return a single **Action** object from the **Actions** collection object of an Outlook item, such as **[MailItem](Outlook.MailItem.md)**.


## Example

The following Visual Basic for Applications (VBA) example uses the Reply action of a particular item to send a reply.


```vb
myItem = CreateItem(olMailItem) 
 
Set myReply = myItem.Actions("Reply").Execute
```

The following Visual Basic for Applications example does the same thing, using a different reply style for the reply.




```vb
myItem = CreateItem(olMailItem) 
 
myItem.Actions("Reply").ReplyStyle = _ 
 
 olIncludeOriginalText 
 
Set myReply = myItem.Actions("Reply").Execute
```


## Methods



|Name|
|:-----|
|[Delete](Outlook.Action.Delete.md)|
|[Execute](Outlook.Action.Execute.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Action.Application.md)|
|[Class](Outlook.Action.Class.md)|
|[CopyLike](Outlook.Action.CopyLike.md)|
|[Enabled](Outlook.Action.Enabled.md)|
|[MessageClass](Outlook.Action.MessageClass.md)|
|[Name](Outlook.Action.Name.md)|
|[Parent](Outlook.Action.Parent.md)|
|[Prefix](Outlook.Action.Prefix.md)|
|[ReplyStyle](Outlook.Action.ReplyStyle.md)|
|[ResponseStyle](Outlook.Action.ResponseStyle.md)|
|[Session](Outlook.Action.Session.md)|
|[ShowOn](Outlook.Action.ShowOn.md)|

## See also


[Action Object Members](overview/Outlook.md)
[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]