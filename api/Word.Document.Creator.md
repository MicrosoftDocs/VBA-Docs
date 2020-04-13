---
title: Document.Creator property (Word)
keywords: vbawd10.chm158008297
f1_keywords:
- vbawd10.chm158008297
ms.prod: word
api_name:
- Word.Document.Creator
ms.assetid: 0ed9cf75-8bae-ba10-4ba0-12a73ff84c08
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]