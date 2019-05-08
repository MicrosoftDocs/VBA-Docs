---
title: Assignment.AppendNotes method (Project)
ms.prod: project-server
api_name:
- Project.Assignment.AppendNotes
ms.assetid: 78ccad76-ac3f-c11e-9d88-2ed133358671
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment.AppendNotes method (Project)

Appends text to the Notes field.


## Syntax

_expression_. `AppendNotes`( `_Value_` )

_expression_ A variable that represents an [Assignment](./Project.Assignment.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]