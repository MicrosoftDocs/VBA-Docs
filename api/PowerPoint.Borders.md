---
title: Borders object (PowerPoint)
keywords: vbapp10.chm629000
f1_keywords:
- vbapp10.chm629000
ms.prod: powerpoint
api_name:
- PowerPoint.Borders
ms.assetid: af3b8d8b-9214-b1ac-f12e-0be244b60b08
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders object (PowerPoint)

A collection of  **[LineFormat](PowerPoint.LineFormat.md)** objects that represent the borders and diagonal lines of a cell or range of cells in a table.


## Remarks

Each  **Cell** object or **CellRange** collection has six elements in the **Borders** collection. You cannot add objects to the **Borders** collection.

Use  **Borders** (index), where index identifies the cell border or diagonal line, to return a single **Border** object. The index value can be any **PPBorderType** constant.


||
|:-----|
|**ppBorderBottom**|
|**ppBorderLeft**|
|**ppBorderRight**|
|**ppBorderTop**|
|**ppBorderDiagonalDown**|
|**ppBorderDiagonalUp**|

## Example

Use the [DashStyle](PowerPoint.LineFormat.DashStyle.md)property to apply a dashed line style to a **Border** object. This example selects the second row from the table and applies a dashed line style to the bottom border.


```vb
ActiveWindow.Selection.ShapeRange.Table.Rows(2).Cells.Borders(ppBorderBottom).DashStyle = msoLineDash
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]