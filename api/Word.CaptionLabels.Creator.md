---
title: CaptionLabels.Creator property (Word)
keywords: vbawd10.chm158860265
f1_keywords:
- vbawd10.chm158860265
ms.prod: word
api_name:
- Word.CaptionLabels.Creator
ms.assetid: 956eec82-9d92-880c-83ad-2437e7bc6e41
ms.date: 06/08/2017
localization_priority: Normal
---


# CaptionLabels.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[CaptionLabels](Word.captionlabels.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[CaptionLabels Collection Object](Word.captionlabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]