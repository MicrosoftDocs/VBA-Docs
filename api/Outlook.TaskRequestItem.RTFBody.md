---
title: TaskRequestItem.RTFBody property (Outlook)
keywords: vbaol11.chm3531
f1_keywords:
- vbaol11.chm3531
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.RTFBody
ms.assetid: c5bea0fa-02e2-20ab-0d81-541478cfd1f0
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.RTFBody property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

_expression_. `RTFBody`

_expression_ A variable that represents a '[TaskRequestItem](Outlook.TaskRequestItem.md)' object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]