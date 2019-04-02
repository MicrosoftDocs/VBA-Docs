---
title: OlActionReplyStyle enumeration (Outlook)
keywords: vbaol11.chm3049
f1_keywords:
- vbaol11.chm3049
ms.prod: outlook
api_name:
- Outlook.OlActionReplyStyle
ms.assetid: 730f9712-a2bb-f698-d210-9dc94da373e8
ms.date: 06/08/2017
localization_priority: Normal
---


# OlActionReplyStyle enumeration (Outlook)

Specifies the reply style.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olEmbedOriginalItem**|1|The reply will include the original item embedded in it. |
| **olIncludeOriginalText**|2|The reply will include the text of the original item.|
| **olIndentOriginalText**|3|The reply will include the indented text of the original item.|
| **olLinkOriginalItem**|4|The reply will include a link to the original item.|
| **olOmitOriginalText**|0|The reply will not include any references to the original item or its text.|
| **olReplyTickOriginalText**|1000|The reply will include the original text with each line preceded by a symbol such as ">".|
| **olUserPreference**|5|The reply style will be set based on the user's preference.|

## Remarks

Used by the [ReplyStyle](Outlook.Action.ReplyStyle.md) property of an [Action](Outlook.Action.md) to specify the reply style that will be used when the **Action** is executed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]