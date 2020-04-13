---
title: TaskRequestItem.Move method (Outlook)
keywords: vbaol11.chm1902
f1_keywords:
- vbaol11.chm1902
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Move
ms.assetid: 9a33da92-aa10-fe5a-b5d2-9c68be1886e5
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.Move method (Outlook)

Moves a Microsoft Outlook item to a new folder.


## Syntax

_expression_. `Move`( `_DestFldr_` )

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DestFldr_|Required| **[Folder](Outlook.Folder.md)**|An expression that returns a **Folder** object. The destination folder.|

## Return value

An **Object** value that represents the item which has been moved to the designated folder.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]