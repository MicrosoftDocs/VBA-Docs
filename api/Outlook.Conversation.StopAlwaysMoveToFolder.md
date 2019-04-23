---
title: Conversation.StopAlwaysMoveToFolder method (Outlook)
keywords: vbaol11.chm3433
f1_keywords:
- vbaol11.chm3433
ms.prod: outlook
api_name:
- Outlook.Conversation.StopAlwaysMoveToFolder
ms.assetid: 3be830e9-ceea-369c-1f7b-966c68cfb8fd
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.StopAlwaysMoveToFolder method (Outlook)

Stops the action of always moving conversation items in the specified store to a specific folder.


## Syntax

_expression_. `StopAlwaysMoveToFolder`( `_Store_` )

_expression_ A variable that represents a '[Conversation](Outlook.Conversation.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](Outlook.Store.md)**|The store where the conversation items to be cleaned up reside.|

## Remarks

If the always-move action has not been turned on,  **StopAlwaysMoveToFolder** does not carry out any action.

If the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the stop-always-move action will apply to conversation items in the default delivery store.

After you call the  **StopAlwaysMoveToFolder** method, calling the **[GetAlwaysMoveToFolder](Outlook.Conversation.GetAlwaysMoveToFolder.md)** method returns **Null** (**Nothing** in Visual Basic).


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]