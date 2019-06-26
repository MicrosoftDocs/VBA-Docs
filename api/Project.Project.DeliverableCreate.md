---
title: Project.DeliverableCreate method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableCreate
ms.assetid: 538f8143-0c0d-b9fa-9219-5405f4bd5046
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.DeliverableCreate method (Project)

Creates a deliverable for a published project that has a project workspace.


## Syntax

_expression_. `DeliverableCreate`( `_DeliverableName_`, `_DeliverableStartDate_`, `_DeliverableFinishDate_`, `_TaskGuid_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DeliverableName_|Required|**String**|Name of the deliverable.|
| _DeliverableStartDate_|Required|**Variant**|Start date for the deliverable.|
| _DeliverableFinishDate_|Required|**Variant**|Finish date for the deliverable.|
| _TaskGuid_|Required|**String**|GUID of the task to which to link the deliverable.|

## Return value

 **String**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]