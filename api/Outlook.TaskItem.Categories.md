---
title: TaskItem.Categories property (Outlook)
keywords: vbaol11.chm1690
f1_keywords:
- vbaol11.chm1690
ms.prod: outlook
api_name:
- Outlook.TaskItem.Categories
ms.assetid: c4099fe0-23af-a4cb-dfef-92cbe0c6e600
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Categories property (Outlook)

Returns or sets a **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]