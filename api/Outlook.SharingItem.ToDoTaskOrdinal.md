---
title: SharingItem.ToDoTaskOrdinal property (Outlook)
keywords: vbaol11.chm3222
f1_keywords:
- vbaol11.chm3222
ms.prod: outlook
api_name:
- Outlook.SharingItem.ToDoTaskOrdinal
ms.assetid: 4164fa78-c0cf-e359-2707-025d6d49f145
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.ToDoTaskOrdinal property (Outlook)

Returns or sets a  **Date** value that represents the ordinal value of the task for the **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `ToDoTaskOrdinal`

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if the **[IsMarkedAsTask](Outlook.SharingItem.IsMarkedAsTask.md)** property is set to **False**.

This property is used to indicate how the task should be ordered within the parent groups, such as the  **Today** group or the **Tomorrow** group, of the **To-Do Bar**. The value used in this property does not have any relation to the values of the  **[TaskStartDate](Outlook.SharingItem.TaskStartDate.md)**, **[TaskDueDate](Outlook.SharingItem.TaskDueDate.md)**, or **[TaskCompletedDate](Outlook.SharingItem.TaskCompletedDate.md)** properties.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]