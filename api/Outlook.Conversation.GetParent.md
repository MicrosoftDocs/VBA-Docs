---
title: Conversation.GetParent method (Outlook)
keywords: vbaol11.chm3401
f1_keywords:
- vbaol11.chm3401
ms.prod: outlook
api_name:
- Outlook.Conversation.GetParent
ms.assetid: edcd31fb-f62e-4273-f827-ac1f704adc5e
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.GetParent method (Outlook)

Returns the parent item of the specified node in the conversation.


## Syntax

_expression_. `GetParent`( `_Item_` )

_expression_ A variable that represents a '[Conversation](Outlook.Conversation.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|A conversation node that is part of a conversation.|

## Return value

An **Object** that represents the parent conversation item of the node specified by the _Item_ parameter.


## Remarks

If the node specified by the  _Item_ parameter does not exist in the conversation, the **GetParent** method returns an error.

If the node specified by the  _Item_ parameter does not have a parent item in the conversation, the **GetParent** method returns **Null** (**Nothing** in Visual Basic).


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]