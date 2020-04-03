---
title: DistListItem.RTFBody property (Outlook)
keywords: vbaol11.chm3529
f1_keywords:
- vbaol11.chm3529
ms.prod: outlook
api_name:
- Outlook.DistListItem.RTFBody
ms.assetid: 0ae5956c-df1e-9ef4-116e-869b69fc11e6
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.RTFBody property (Outlook)

Returns or sets a **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

_expression_. `RTFBody`

_expression_ A variable that represents a '[DistListItem](Outlook.DistListItem.md)' object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]