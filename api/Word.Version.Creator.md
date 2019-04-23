---
title: Version.Creator property (Word)
keywords: vbawd10.chm162792425
f1_keywords:
- vbawd10.chm162792425
ms.prod: word
api_name:
- Word.Version.Creator
ms.assetid: 952bc69e-de7d-faaa-d598-7d1e493ea9a4
ms.date: 06/08/2017
localization_priority: Normal
---


# Version.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Version](Word.Version.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Version Object](Word.Version.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]