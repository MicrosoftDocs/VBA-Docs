---
title: ReportItem.ReplyAll event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.ReplyAll
ms.assetid: b5724798-8c73-13ce-23d4-9d7ec8147f44
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.ReplyAll event (Outlook)

Occurs when the user selects the  **ReplyAll** action for an item (which is an instance of the parent object).


## Syntax

_expression_. `ReplyAll`( `_Response_` , `_Cancel_` )

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Response_|Required| **Object**|The new item being sent in response to the original message.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the reply all operation is not completed and the new item is not displayed.|

## Remarks

Returns the reply as a **[MailItem](Outlook.MailItem.md)** object.


## See also


[ReportItem Object](Outlook.ReportItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]