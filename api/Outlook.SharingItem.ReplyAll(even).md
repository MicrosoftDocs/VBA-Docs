---
title: SharingItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.ReplyAll
ms.assetid: 147f7da9-fa4b-b678-f600-25a8c6b540ec
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item, or when the **[ReplyAll](Outlook.SharingItem.ReplyAll(method).md)** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a  **[MailItem](Outlook.MailItem.md)** object.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]