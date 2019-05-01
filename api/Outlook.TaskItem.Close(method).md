---
title: TaskItem.Close method (Outlook)
keywords: vbaol11.chm1712
f1_keywords:
- vbaol11.chm1712
ms.prod: outlook
api_name:
- Outlook.TaskItem.Close
ms.assetid: 7682f0c8-d132-2bd6-94e8-6e45fcc00867
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.Close method (Outlook)

Closes and optionally saves changes to the Outlook item.


## Syntax

_expression_.**Close** (_SaveMode_)

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](Outlook.OlInspectorClose.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]