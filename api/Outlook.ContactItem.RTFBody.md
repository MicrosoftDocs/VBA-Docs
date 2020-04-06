---
title: ContactItem.RTFBody property (Outlook)
keywords: vbaol11.chm3525
f1_keywords:
- vbaol11.chm3525
ms.prod: outlook
api_name:
- Outlook.ContactItem.RTFBody
ms.assetid: f8e7e632-113b-a50e-211b-dbd182221168
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.RTFBody property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

_expression_. `RTFBody`

_expression_ A variable that represents a '[ContactItem](Outlook.ContactItem.md)' object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]