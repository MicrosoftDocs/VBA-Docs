---
title: CheckBox.Creator property (Word)
keywords: vbawd10.chm153486313
f1_keywords:
- vbawd10.chm153486313
ms.prod: word
api_name:
- Word.CheckBox.Creator
ms.assetid: f06ad900-28b8-2823-0c6a-c535fcae6a4f
ms.date: 06/08/2017
localization_priority: Normal
---


# CheckBox.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[CheckBox](Word.CheckBox.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[CheckBox Object](Word.CheckBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]