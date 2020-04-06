---
title: Actions object (Outlook)
keywords: vbaol11.chm144
f1_keywords:
- vbaol11.chm144
ms.prod: outlook
api_name:
- Outlook.Actions
ms.assetid: b0903aa4-9b75-5311-d0a5-5ff4a5e29c79
ms.date: 06/08/2017
localization_priority: Normal
---


# Actions object (Outlook)

Contains a collection of  **[Action](Outlook.Action.md)** objects that represent all the specialized actions that can be executed on an Outlook item.


## Remarks

Use the  **Actions** property of any Outlook item, such as **[MailItem](Outlook.MailItem.md)**, to return the **Actions** object.

Use  **Actions** (_index_), where _index_ is the name of an available action, to return a single **Action** object.


## Example

The following Visual Basic for Applications (VBA) example uses the Reply action of a particular item to send a reply.


```vb
myItem = CreateItem(olMailItem) 
 
Set myReply = myItem.Actions("Reply").Execute
```


## Methods



|Name|
|:-----|
|[Add](Outlook.Actions.Add.md)|
|[Item](Outlook.Actions.Item.md)|
|[Remove](Outlook.Actions.Remove.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Actions.Application.md)|
|[Class](Outlook.Actions.Class.md)|
|[Count](Outlook.Actions.Count.md)|
|[Parent](Outlook.Actions.Parent.md)|
|[Session](Outlook.Actions.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Actions Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]