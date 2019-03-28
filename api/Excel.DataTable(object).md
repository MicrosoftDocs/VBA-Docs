---
title: DataTable object (Excel)
keywords: vbaxl10.chm625072
f1_keywords:
- vbaxl10.chm625072
ms.prod: excel
api_name:
- Excel.DataTable
ms.assetid: aca0850b-2e72-cde9-b751-633876e1df99
ms.date: 03/29/2019
localization_priority: Normal
---


# DataTable object (Excel)

Represents a chart data table.


## Example

Use the **[DataTable](Excel.Chart.DataTable.md)** property of the **Chart** object to return a **DataTable** object. The following example adds a data table with an outline border to embedded chart one.

```vb
With Worksheets(1).ChartObjects(1).Chart 
 .HasDataTable = True 
 .DataTable.HasBorderOutline = True 
End With
```


## Methods

- [Delete](Excel.DataTable.Delete.md)
- [Select](Excel.DataTable.Select.md)

## Properties

- [Application](Excel.DataTable.Application.md)
- [Border](Excel.DataTable.Border.md)
- [Creator](Excel.DataTable.Creator.md)
- [Font](Excel.DataTable.Font.md)
- [Format](Excel.DataTable.Format.md)
- [HasBorderHorizontal](Excel.DataTable.HasBorderHorizontal.md)
- [HasBorderOutline](Excel.DataTable.HasBorderOutline.md)
- [HasBorderVertical](Excel.DataTable.HasBorderVertical.md)
- [Parent](Excel.DataTable.Parent.md)
- [ShowLegendKey](Excel.DataTable.ShowLegendKey.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
