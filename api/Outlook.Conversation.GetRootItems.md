---
title: Conversation.GetRootItems method (Outlook)
keywords: vbaol11.chm3402
f1_keywords:
- vbaol11.chm3402
ms.prod: outlook
api_name:
- Outlook.Conversation.GetRootItems
ms.assetid: 72c4d9fd-4f38-d081-7dc6-e9dbfad6d3aa
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.GetRootItems method (Outlook)

Returns a  **[SimpleItems](Outlook.SimpleItems.md)** collection that contains all root items in the conversation.


## Syntax

_expression_. `GetRootItems`

_expression_ A variable that represents a '[Conversation](Outlook.Conversation.md)' object.


## Return value

A  **SimpleItems** collection that includes the root item or all root items of the conversation.


## Remarks

A conversation can have one or more root items. For example, if the root item of the conversation has three child items and the root item is permanently deleted, all three child items become root items.

If all items are deleted from the conversation after the  **[Conversation](Outlook.Conversation.md)** object has been obtained, **GetRootItems** returns a **SimpleItems** collection with zero objects. In this case, the **[Count](Outlook.SimpleItems.Count.md)** property of the **SimpleItems** collection returns 0.


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]