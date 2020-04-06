---
title: DistListItem.ToDoTaskOrdinal property (Outlook)
keywords: vbaol11.chm3034
f1_keywords:
- vbaol11.chm3034
ms.prod: outlook
api_name:
- Outlook.DistListItem.ToDoTaskOrdinal
ms.assetid: a72f8ba4-a31a-b96b-193a-2507b7c45169
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.ToDoTaskOrdinal property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[DistListItem](Outlook.DistListItem.md)**. Read/write.


## Syntax

_expression_. `ToDoTaskOrdinal`

 _expression_ An expression that returns a [DistListItem](Outlook.DistListItem.md) object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if the **[IsMarkedAsTask](Outlook.DistListItem.IsMarkedAsTask.md)** property is set to **False**.

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](Outlook.DistListItem.TaskStartDate.md)**, **[TaskDueDate](Outlook.DistListItem.TaskDueDate.md)**, or **[TaskCompletedDate](Outlook.DistListItem.TaskCompletedDate.md)** properties.


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]