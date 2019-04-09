---
title: TableStyle.Creator property (Word)
keywords: vbawd10.chm244777961
f1_keywords:
- vbawd10.chm244777961
ms.prod: word
api_name:
- Word.TableStyle.Creator
ms.assetid: 995f01de-80dd-6c57-432d-24c04ad7d1f0
ms.date: 06/08/2017
localization_priority: Normal
---


# TableStyle.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[TableStyle](Word.TableStyle.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[TableStyle Object](Word.TableStyle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]