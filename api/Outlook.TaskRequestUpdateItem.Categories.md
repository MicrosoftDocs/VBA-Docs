---
title: TaskRequestUpdateItem.Categories property (Outlook)
keywords: vbaol11.chm1925
f1_keywords:
- vbaol11.chm1925
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Categories
ms.assetid: a4e0c824-fc22-76b0-e9e5-03265aec7066
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.Categories property (Outlook)

Returns or sets a  **String** representing the categories assigned to the Outlook item. Read/write.


## Syntax

_expression_. `Categories`

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Remarks

 **Categories** is a delimited string of category names that have been assigned to an Outlook item. This property uses the character specified in the value name, **sList**, under **HKEY_CURRENT_USER\Control Panel\International** in the Windows registry, as the delimiter for multiple categories. To convert the string of category names to an array of category names, use the Microsoft Visual Basic function **Split**.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]