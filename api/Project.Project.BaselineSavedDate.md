---
title: Project.BaselineSavedDate property (Project)
ms.prod: project-server
api_name:
- Project.Project.BaselineSavedDate
ms.assetid: 780c5190-68bb-1c10-0dbb-612e5606184e
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.BaselineSavedDate property (Project)

Gets date the specified baseline was last saved. Read-only  **Variant**.


## Syntax

_expression_. `BaselineSavedDate`( `_Baseline_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Baseline_|Required|**Long**|Can be one of the **[PjBaselines](Project.PjBaselines.md)** constants.|

## Remarks

If the specified baseline has not been saved,  **BaselineSavedDate** returns "NA".

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]