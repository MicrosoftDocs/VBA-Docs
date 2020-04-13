---
title: Breaks.Creator property (Word)
keywords: vbawd10.chm50922473
f1_keywords:
- vbawd10.chm50922473
ms.prod: word
api_name:
- Word.Breaks.Creator
ms.assetid: 4324547c-df85-62ed-c238-74b9b9b6deeb
ms.date: 06/08/2017
localization_priority: Normal
---


# Breaks.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an '[Breaks](Word.Breaks.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Breaks Collection](Word.Breaks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]