---
title: Resource.AppendNotes method (Project)
ms.prod: project-server
api_name:
- Project.Resource.AppendNotes
ms.assetid: b11bc28f-147f-0591-056b-87e9f6c2db71
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.AppendNotes method (Project)

Appends text to the Notes field.


## Syntax

_expression_. `AppendNotes`( `_Value_` )

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]