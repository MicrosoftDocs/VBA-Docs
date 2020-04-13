---
title: MappedDataField.Creator property (Word)
keywords: vbawd10.chm107545577
f1_keywords:
- vbawd10.chm107545577
ms.prod: word
api_name:
- Word.MappedDataField.Creator
ms.assetid: d4c432ea-9b14-8c03-2f0c-36e25e58d23f
ms.date: 06/08/2017
localization_priority: Normal
---


# MappedDataField.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[MappedDataField](Word.MappedDataField.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[MappedDataField Object](Word.MappedDataField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]