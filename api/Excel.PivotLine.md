---
title: PivotLine object (Excel)
keywords: vbaxl10.chm763072
f1_keywords:
- vbaxl10.chm763072
ms.prod: excel
api_name:
- Excel.PivotLine
ms.assetid: 88961b73-2d9f-1112-5dd5-14c1fa02092f
ms.date: 03/30/2019
localization_priority: Normal
---


# PivotLine object (Excel)

A **PivotLine** object is a line of rows or columns in an Excel PivotTable.


## Remarks

PivotLines contain only visible items, so collapsed children of items and items in hidden levels are not present in the **[PivotLines](excel.pivotlines.md)** collection.

PivotLines always have a PivotItem in all positions. This means that the PivotLines representing subtotals in the PivotTable contain fewer PivotItems than regular PivotLines.

## Properties

- [Application](Excel.PivotLine.Application.md)
- [Creator](Excel.PivotLine.Creator.md)
- [LineType](Excel.PivotLine.LineType.md)
- [Parent](Excel.PivotLine.Parent.md)
- [PivotLineCells](Excel.PivotLine.PivotLineCells.md)
- [PivotLineCellsFull](Excel.pivotline.pivotlinecellsfull.md)
- [Position](Excel.PivotLine.Position.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]