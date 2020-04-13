---
title: Subdocuments.Creator property (Word)
keywords: vbawd10.chm159908841
f1_keywords:
- vbawd10.chm159908841
ms.prod: word
api_name:
- Word.Subdocuments.Creator
ms.assetid: 6419c182-d35b-eafa-8f1d-17fe297c6037
ms.date: 06/08/2017
localization_priority: Normal
---


# Subdocuments.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Subdocuments](Word.subdocuments.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Subdocuments Collection Object](Word.subdocuments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]