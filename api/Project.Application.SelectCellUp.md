---
title: Application.SelectCellUp method (Project)
keywords: vbapj.chm2049
f1_keywords:
- vbapj.chm2049
ms.prod: project-server
api_name:
- Project.Application.SelectCellUp
ms.assetid: d2e2aecc-0a05-7dd5-23da-a47ffe161028
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SelectCellUp method (Project)

Selects cells upward from the current selection.


## Syntax

_expression_. `SelectCellUp`( `_NumCells_`, `_Extend_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumCells_|Optional|**Long**|The number of cells to select upward from the current selection. The default value is 1.|
| _Extend_|Optional|**Boolean**|**True** if the current selection is extended to the specified cell. The default value is **False**.|

## Return value

 **Boolean**


## Remarks

The  **SelectCellUp** method is not available when the Calendar, Network Diagram, or Resource Graph is the active view.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]