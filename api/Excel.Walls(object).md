---
title: Walls object (Excel)
keywords: vbaxl10.chm613072
f1_keywords:
- vbaxl10.chm613072
ms.prod: excel
api_name:
- Excel.Walls
ms.assetid: 9c6f0c5b-dbb8-7d71-44b7-29987e750cd3
ms.date: 04/03/2019
localization_priority: Normal
---


# Walls object (Excel)

Represents the walls of a 3D chart. This object isn't a collection. There's no object that represents a single wall; you must return all the walls as a unit.


## Example

Use the **[Walls](Excel.Chart.Walls.md)** property of the **Chart** object to return the **Walls** object. 

The following example sets the pattern on the walls for embedded chart one on Sheet1. If the chart isn't a 3D chart, this example fails.

```vb
Worksheets("Sheet1").ChartObjects(1).Chart _ 
 .Walls.Interior.Pattern = xlGray75
```

## Methods

- [ClearFormats](Excel.Walls.ClearFormats.md)
- [Paste](Excel.Walls.Paste.md)
- [Select](Excel.Walls.Select.md)

## Properties

- [Application](Excel.Walls.Application.md)
- [Creator](Excel.Walls.Creator.md)
- [Format](Excel.Walls.Format.md)
- [Name](Excel.Walls.Name.md)
- [Parent](Excel.Walls.Parent.md)
- [PictureType](Excel.Walls.PictureType.md)
- [PictureUnit](Excel.Walls.PictureUnit.md)
- [Thickness](Excel.Walls.Thickness.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]