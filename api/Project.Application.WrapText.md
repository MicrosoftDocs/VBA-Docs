---
title: Application.WrapText method (Project)
keywords: vbapj.chm708
f1_keywords:
- vbapj.chm708
ms.prod: project-server
api_name:
- Project.Application.WrapText
ms.assetid: 0aaabac2-ee1d-694c-45ac-f522a0034724
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WrapText method (Project)

Toggles the **Wrap Text** setting in a column.


## Syntax

_expression_.**WrapText** (_Column_)

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**Integer**|The target column identifier. If omitted, the **WrapText** method is applied to the column containing the active cell.|

## Return value

**Boolean**


## Remarks

The **WrapText** method corresponds to the **Wrap Text** command in the option menu for a column.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]