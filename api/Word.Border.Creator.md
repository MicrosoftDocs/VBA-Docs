---
title: Border.Creator property (Word)
keywords: vbawd10.chm154862569
f1_keywords:
- vbawd10.chm154862569
ms.prod: word
api_name:
- Word.Border.Creator
ms.assetid: 3e372111-4449-b3ef-e572-3cb0db4dcc69
ms.date: 06/08/2017
localization_priority: Normal
---


# Border.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an '[Border](Word.Border.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Border Object](Word.Border.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]