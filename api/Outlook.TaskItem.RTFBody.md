---
title: TaskItem.RTFBody property (Outlook)
keywords: vbaol11.chm3528
f1_keywords:
- vbaol11.chm3528
ms.prod: outlook
api_name:
- Outlook.TaskItem.RTFBody
ms.assetid: ff94ab2c-7e34-0eb5-3aeb-b7805b5e9a2c
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.RTFBody property (Outlook)

Returns or sets a **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

_expression_. `RTFBody`

_expression_ A variable that represents a '[TaskItem](Outlook.TaskItem.md)' object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]