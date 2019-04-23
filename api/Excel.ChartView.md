---
title: ChartView object (Excel)
keywords: vbaxl10.chm780072
f1_keywords:
- vbaxl10.chm780072
ms.prod: excel
api_name:
- Excel.ChartView
ms.assetid: 2e59e8c1-f1cd-1589-ae36-22d6c5dccbf6
ms.date: 03/29/2019
localization_priority: Normal
---


# ChartView object (Excel)

Represents a view of a chart.


## Remarks

The **ChartView** object is one of the objects that can be returned by the **[SheetViews](Excel.SheetViews.md)** collection, similar to the **[Sheets](Excel.Sheets.md)** collection. The **ChartView** object applies only to chart sheets.


## Example

The following example returns a **ChartView** object.

```vb
ActiveWindow.SheetViews.Item(1) 

```

<br/>

The following example returns a **[Chart](Excel.Chart(object).md)** object.

```vb
ActiveWindow.SheetViews.Item(1).Sheet 

```


## Properties

- [Application](Excel.ChartView.Application.md)
- [Creator](Excel.ChartView.Creator.md)
- [Parent](Excel.ChartView.Parent.md)
- [Sheet](Excel.ChartView.Sheet.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]