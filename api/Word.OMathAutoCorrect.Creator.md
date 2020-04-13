---
title: OMathAutoCorrect.Creator property (Word)
keywords: vbawd10.chm258474085
f1_keywords:
- vbawd10.chm258474085
ms.prod: word
api_name:
- Word.OMathAutoCorrect.Creator
ms.assetid: 2f2fcc01-799b-5764-4997-fe49113fb2cd
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathAutoCorrect.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[OMathAutoCorrect](Word.OMathAutoCorrect.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[OMathAutoCorrect Object](Word.OMathAutoCorrect.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]