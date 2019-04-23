---
title: TaskItem.Move method (Outlook)
keywords: vbaol11.chm1716
f1_keywords:
- vbaol11.chm1716
ms.prod: outlook
api_name:
- Outlook.TaskItem.Move
ms.assetid: cc071e73-d165-6082-4016-7ab9d63689d0
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Move method (Outlook)

Moves a Microsoft Outlook item to a new folder.


## Syntax

_expression_. `Move`( `_DestFldr_` )

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DestFldr_|Required| **[Folder](Outlook.Folder.md)**|An expression that returns a  **Folder** object. The destination folder.|

## Return value

An  **Object** value that represents the item which has been moved to the designated folder.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]