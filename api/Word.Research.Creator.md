---
title: Research.Creator property (Word)
ms.prod: word
api_name:
- Word.Research.Creator
ms.assetid: 5947e75d-97b3-0d6a-9241-1843ab76c635
ms.date: 06/08/2017
localization_priority: Normal
---


# Research.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[Research](Word.Research.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Research Object](Word.Research.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]