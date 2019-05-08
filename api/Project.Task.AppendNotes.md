---
title: Task.AppendNotes method (Project)
ms.prod: project-server
api_name:
- Project.Task.AppendNotes
ms.assetid: ab0177cb-c7cd-444f-0d19-9b798eba8b4a
ms.date: 06/08/2017
localization_priority: Normal
---


# Task.AppendNotes method (Project)

Appends text to the Notes field.


## Syntax

_expression_. `AppendNotes`( `_Value_` )

_expression_ A variable that represents a [Task](./Project.Task.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**String**|The text to append to the existing notes.|

## Remarks

New text is added with the formatting in use at the end of any existing notes.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]