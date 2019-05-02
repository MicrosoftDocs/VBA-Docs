---
title: TaskRequestItem.Close method (Outlook)
keywords: vbaol11.chm1898
f1_keywords:
- vbaol11.chm1898
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.Close
ms.assetid: c24b364b-f4d5-22dc-2357-691311e9f34b
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.Close method (Outlook)

Closes and optionally saves changes to the Outlook item.


## Syntax

_expression_.**Close** (_SaveMode_)

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](Outlook.OlInspectorClose.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]