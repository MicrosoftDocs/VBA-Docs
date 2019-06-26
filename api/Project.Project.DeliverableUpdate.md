---
title: Project.DeliverableUpdate method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableUpdate
ms.assetid: 665e79a0-b3b4-e36e-6369-627e526f7db0
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.DeliverableUpdate method (Project)

Updates the properties of a deliverable.


## Syntax

_expression_. `DeliverableUpdate`( `_DeliverableGuid_`, `_DeliverableName_`, `_DeliverableStartDate_`, `_DeliverableFinishDate_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DeliverableGuid_|Required|**String**|GUID of the deliverable to update.|
| _DeliverableName_|Required|**String**|Name of the deliverable.|
| _DeliverableStartDate_|Required|**Variant**|Date when the deliverable starts.|
| _DeliverableFinishDate_|Required|**Variant**|Date when the deliverable is finished.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]