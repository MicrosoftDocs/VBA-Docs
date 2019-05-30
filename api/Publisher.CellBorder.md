---
title: CellBorder object (Publisher)
keywords: vbapb10.chm5308415
f1_keywords:
- vbapb10.chm5308415
ms.prod: publisher
api_name:
- Publisher.CellBorder
ms.assetid: c4eddeac-54cd-95ff-9423-b06e515a720e
ms.date: 05/31/2019
localization_priority: Normal
---


# CellBorder object (Publisher)

Represents the color and weight settings for cell borders.
 
## Remarks

Use the various border properties of the **[Cell](Publisher.Cell.md)** object to return the different borders of a cell (left, right, top, bottom, and diagonal).

Use the **Color** and **Weight** properties of the **CellBorder** object to format the appearance of a cell border. 

## Example

The following example retrieves the top border of the first cell in a table.

```vb
Dim cbTemp As CellBorder 
 
Set cbTemp = ActiveDocument.Pages(1) _ 
 .Shapes(1).Table.Cells.Item(1).BorderTop
```

<br/>

The following example makes the left border of the first cell in a table red and two points thick.

```vb
Dim cbTemp As CellBorder 
 
Set cbTemp = ActiveDocument.Pages(1) _ 
 .Shapes(1).Table.Cells.Item(1).BorderLeft 
 
cbTemp.Color.RGB = RGB(255, 0, 0) 
cbTemp.Weight = 2
```


## Properties

- [Application](Publisher.CellBorder.Application.md)
- [Color](Publisher.CellBorder.Color.md)
- [Parent](Publisher.CellBorder.Parent.md)
- [Weight](Publisher.CellBorder.Weight.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]