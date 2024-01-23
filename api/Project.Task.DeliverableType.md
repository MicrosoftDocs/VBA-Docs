---
title: Task.DeliverableType property (Project)
ms.service: project-server
api_name:
- Project.Task.DeliverableType
ms.assetid: 4170340d-ea80-54ab-b65a-08ee062ad41b
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Task.DeliverableType property (Project)

Gets or sets the type of deliverable for the task. Read/write **Integer**.


## Syntax

_expression_. `DeliverableType`

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Remarks

The **DeliverableType** property can have the following values:



|Value|Description|
|:-----|:-----|
|0|The task has no associated deliverable.|
|1|The associated deliverable is produced by the task.|
|2|The associated deliverable is produced by a separate project or task upon which the current task is dependent.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]