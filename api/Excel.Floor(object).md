---
title: Floor object (Excel)
keywords: vbaxl10.chm611072
f1_keywords:
- vbaxl10.chm611072
ms.prod: excel
api_name:
- Excel.Floor
ms.assetid: 74c71ca8-a0d4-f7cf-a002-5cec7a27b70d
ms.date: 03/29/2019
localization_priority: Normal
---


# Floor object (Excel)

Represents the floor of a 3D chart.


## Example

Use the **[Floor](Excel.Chart.Floor.md)** property of the **Chart** object to return the **Floor** object. The following example sets the floor color for embedded chart one to cyan. The example will fail if the chart isn't a 3D chart.

```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.Floor.Interior.Color = RGB(0, 255, 255)
```

## Methods

- [ClearFormats](Excel.Floor.ClearFormats.md)
- [Paste](Excel.Floor.Paste.md)
- [Select](Excel.Floor.Select.md)

## Properties

- [Application](Excel.Floor.Application.md)
- [Creator](Excel.Floor.Creator.md)
- [Format](Excel.Floor.Format.md)
- [Name](Excel.Floor.Name.md)
- [Parent](Excel.Floor.Parent.md)
- [PictureType](Excel.Floor.PictureType.md)
- [Thickness](Excel.Floor.Thickness.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]