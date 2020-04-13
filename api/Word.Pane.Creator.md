---
title: Pane.Creator property (Word)
keywords: vbawd10.chm157287401
f1_keywords:
- vbawd10.chm157287401
ms.prod: word
api_name:
- Word.Pane.Creator
ms.assetid: 142cc9bf-41f9-fcb5-73c0-8cc9f21cd553
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Pane](Word.Pane.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]