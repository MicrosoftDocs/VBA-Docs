---
title: OlTaskResponse enumeration (Outlook)
keywords: vbaol11.chm3086
f1_keywords:
- vbaol11.chm3086
ms.prod: outlook
api_name:
- Outlook.OlTaskResponse
ms.assetid: 7616cbdc-fc9c-abbe-fd07-ebdadc13ede2
ms.date: 06/08/2017
localization_priority: Normal
---


# OlTaskResponse enumeration (Outlook)

Indicates the response to a task request.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olTaskAccept**|2|Task accepted.|
| **olTaskAssign**|1|Task reassigned.|
| **olTaskDecline**|3|Task declined.|
| **olTaskSimple**|0|Task is a simple task and cannot be accepted, declined, or assigned. This constant is not a valid parameter to the  **TaskItem.Respond** method.|

## Remarks

Used by the [TaskItem.ResponseState property (Outlook)](Outlook.TaskItem.ResponseState.md) and as a parameter to the [TaskItem.Respond method (Outlook)](Outlook.TaskItem.Respond.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]