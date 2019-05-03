---
title: PivotCell.ServerActions property (Excel)
keywords: vbaxl10.chm692090
f1_keywords:
- vbaxl10.chm692090
ms.prod: excel
ms.assetid: e895f7ee-e636-29b6-9385-2710885cc01c
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.ServerActions property (Excel)

Represents a collection of _actions_ consisting of OLAP-defined actions that can be executed. The actions are specific to PivotTables existing at a worksheet-level. Read-only.


## Syntax

_expression_.**ServerActions**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Remarks

A server action is an optional feature that an OLAP cube administrator can define on a server that uses a cube member or measure as a parameter into a query to obtain details in the cube.


## Property value

**ACTIONS**


## Example

The following code segment executes a server action against a series in a PivotChart.

```vb
ActiveSheet.ChartObjects("Chart 1").Chart.PivotLayout.PivotTable.PivotColumnAxis.PivotLines(index of line ).PivotLineCells(index of cells ).ServerAction("OLAP Action name" ).Execute
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]