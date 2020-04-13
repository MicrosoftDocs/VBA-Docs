---
title: ChartGroup.Creator property (Word)
keywords: vbawd10.chm263454869
f1_keywords:
- vbawd10.chm263454869
ms.prod: word
api_name:
- Word.ChartGroup.Creator
ms.assetid: 6c08be09-c7cd-ab41-3f75-fee9f26f6226
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]