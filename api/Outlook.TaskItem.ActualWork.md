---
title: TaskItem.ActualWork property (Outlook)
keywords: vbaol11.chm1720
f1_keywords:
- vbaol11.chm1720
ms.prod: outlook
api_name:
- Outlook.TaskItem.ActualWork
ms.assetid: d61075da-bd14-bc59-8f72-b9b675c65f08
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.ActualWork property (Outlook)

Returns or sets a **Long** indicating the actual effort spent on the task. Read/write.


## Syntax

_expression_. `ActualWork`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

 **ActualWork** is stored in units of minutes. The **ActualWork** field on the standard task form is bound to the **ActualWork** property; by default the field assumes an 8-hour day and 40-hour week.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]