---
title: Project.DeliverableLinkToProject method (Project)
ms.prod: project-server
api_name:
- Project.Project.DeliverableLinkToProject
ms.assetid: aa78de59-13b2-98f8-45e7-2c40edfaeb25
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.DeliverableLinkToProject method (Project)

Links a deliverable or a dependency to a project.


## Syntax

_expression_. `DeliverableLinkToProject`( `_DeliverableGuid_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DeliverableGuid_|Required|**String**|GUID of the deliverable or the dependency to the link.|

## Return value

 **Boolean**


## Remarks

The  **DeliverableLinkToProject** method unlinks the deliverable or dependency from a task.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]