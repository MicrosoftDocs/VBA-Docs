---
title: OutlineCode.OnlyLeaves property (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode.OnlyLeaves
ms.assetid: cc477127-c784-fdea-53b1-7399d18d6b8b
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlineCode.OnlyLeaves property (Project)

 **True** if only outline code lookup table values without children can be selected. Read/write **Boolean**.


## Syntax

_expression_. `OnlyLeaves`

_expression_ A variable that represents an [OutlineCode](./Project.OutlineCode.md) object.


## Remarks

If there are no values in the outline code lookup table, then  **OnlyLeaves** is **False** and non-writeable. For enterprise text fields with a lookup table, **OnlyLeaves** is always **False** and non-writeable.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]