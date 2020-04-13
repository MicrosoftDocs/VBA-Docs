---
title: StyleSheet.Creator property (Word)
keywords: vbawd10.chm166659049
f1_keywords:
- vbawd10.chm166659049
ms.prod: word
api_name:
- Word.StyleSheet.Creator
ms.assetid: 1bbc4ba2-375a-a599-7297-71cfa68129ea
ms.date: 06/08/2017
localization_priority: Normal
---


# StyleSheet.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[StyleSheet](Word.StyleSheet.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[StyleSheet Object](Word.StyleSheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]