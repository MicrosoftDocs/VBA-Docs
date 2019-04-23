---
title: Template.Creator property (Word)
keywords: vbawd10.chm157942761
f1_keywords:
- vbawd10.chm157942761
ms.prod: word
api_name:
- Word.Template.Creator
ms.assetid: 76329e48-d3a8-334b-4c33-75c6f75f8c43
ms.date: 06/08/2017
localization_priority: Normal
---


# Template.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Template](Word.Template.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Template Object](Word.Template.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]