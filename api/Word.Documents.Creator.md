---
title: Documents.Creator property (Word)
keywords: vbawd10.chm158073833
f1_keywords:
- vbawd10.chm158073833
ms.prod: word
api_name:
- Word.Documents.Creator
ms.assetid: c9acedc3-0e0a-7a16-087f-e5c420cd2e3f
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

 _expression_. `Creator`

 _expression_ Required. A variable that represents a '[Documents](Word.documents.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Documents Collection Object](Word.documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]