---
title: ChartTitle object (Excel)
keywords: vbaxl10.chm562072
f1_keywords:
- vbaxl10.chm562072
ms.prod: excel
api_name:
- Excel.ChartTitle
ms.assetid: e0a10650-66dd-dd33-e9ba-5a5c0f78f2c3
ms.date: 03/29/2019
localization_priority: Normal
---


# ChartTitle object (Excel)

Represents the chart title.


## Remarks

Use the **[ChartTitle](excel.chart.charttitle.md)** property of the **Chart** object to return the **ChartTitle** object.

The **ChartTitle** object doesn't exist and cannot be used unless the **[HasTitle](Excel.Chart.HasTitle.md)** property for the chart is **True**.


## Example

The following example adds a title to embedded chart one on the worksheet named **Sheet1**.

```vb
With Worksheets("sheet1").ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Methods

- [Delete](Excel.ChartTitle.Delete.md)
- [Select](Excel.ChartTitle.Select.md)

## Properties

- [Application](Excel.ChartTitle.Application.md)
- [Caption](Excel.ChartTitle.Caption.md)
- [Characters](Excel.ChartTitle.Characters.md)
- [Creator](Excel.ChartTitle.Creator.md)
- [Format](Excel.ChartTitle.Format.md)
- [Formula](Excel.ChartTitle.Formula.md)
- [FormulaLocal](Excel.ChartTitle.FormulaLocal.md)
- [FormulaR1C1](Excel.ChartTitle.FormulaR1C1.md)
- [FormulaR1C1Local](Excel.ChartTitle.FormulaR1C1Local.md)
- [Height](Excel.ChartTitle.Height.md)
- [HorizontalAlignment](Excel.ChartTitle.HorizontalAlignment.md)
- [IncludeInLayout](Excel.ChartTitle.IncludeInLayout.md)
- [Left](Excel.ChartTitle.Left.md)
- [Name](Excel.ChartTitle.Name.md)
- [Orientation](Excel.ChartTitle.Orientation.md)
- [Parent](Excel.ChartTitle.Parent.md)
- [Position](Excel.ChartTitle.Position.md)
- [ReadingOrder](Excel.ChartTitle.ReadingOrder.md)
- [Shadow](Excel.ChartTitle.Shadow.md)
- [Text](Excel.ChartTitle.Text.md)
- [Top](Excel.ChartTitle.Top.md)
- [VerticalAlignment](Excel.ChartTitle.VerticalAlignment.md)
- [Width](Excel.ChartTitle.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
