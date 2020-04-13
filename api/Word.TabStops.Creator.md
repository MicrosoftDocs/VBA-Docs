---
title: TabStops.Creator property (Word)
keywords: vbawd10.chm156566505
f1_keywords:
- vbawd10.chm156566505
ms.prod: word
api_name:
- Word.TabStops.Creator
ms.assetid: 656e7362-a08e-ef90-2996-17ed7bcead6b
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStops.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[TabStops](Word.tabstops.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[TabStops Collection Object](Word.tabstops.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]