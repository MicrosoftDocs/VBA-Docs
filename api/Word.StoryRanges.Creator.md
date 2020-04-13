---
title: StoryRanges.Creator property (Word)
keywords: vbawd10.chm160170985
f1_keywords:
- vbawd10.chm160170985
ms.prod: word
api_name:
- Word.StoryRanges.Creator
ms.assetid: 192e5457-6ef6-4442-708e-5bd3dc96f843
ms.date: 06/08/2017
localization_priority: Normal
---


# StoryRanges.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[StoryRanges](Word.storyranges.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[StoryRanges Collection Object](Word.storyranges.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]