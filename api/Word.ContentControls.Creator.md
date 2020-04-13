---
title: ContentControls.Creator property (Word)
keywords: vbawd10.chm157745253
f1_keywords:
- vbawd10.chm157745253
ms.prod: word
api_name:
- Word.ContentControls.Creator
ms.assetid: c5230afe-4d34-1215-82fc-401fd7d2a372
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControls.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[ContentControls](Word.ContentControls.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[ContentControls Collection](Word.ContentControls.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]