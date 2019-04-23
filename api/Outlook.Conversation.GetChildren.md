---
title: Conversation.GetChildren method (Outlook)
keywords: vbaol11.chm3391
f1_keywords:
- vbaol11.chm3391
ms.prod: outlook
api_name:
- Outlook.Conversation.GetChildren
ms.assetid: bc68ccd6-9d3c-c404-72b0-a21dbc99ed63
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.GetChildren method (Outlook)

Returns a  **[SimpleItems](Outlook.SimpleItems.md)** collection that contains all items under the specified conversation node.


## Syntax

_expression_. `GetChildren`( `_Item_` )

_expression_ A variable that represents a '[Conversation](Outlook.Conversation.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|A conversation node that is part of a conversation.|

## Return value

A  **SimpleItems** collection that represents the set of child items in the conversation under the node specified by the _Item_ parameter.


## Remarks

The returned  **SimpleItems** collection contains immediate child items of the conversation node specified by the _Item_ parameter. If the specified node does not exist in the conversation, the **GetChildren** method returns an error.

If no child items exist under that node, the  **GetChildren** method returns a **SimpleItems** collection with zero objects, in which case the **[Count](Outlook.SimpleItems.Count.md)** property of the **SimpleItems** collection returns 0.


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]