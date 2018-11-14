---
title: ChartView object (Excel)
keywords: vbaxl10.chm780072
f1_keywords:
- vbaxl10.chm780072
ms.prod: excel
api_name:
- Excel.ChartView
ms.assetid: 2e59e8c1-f1cd-1589-ae36-22d6c5dccbf6
ms.date: 06/08/2017
---


# ChartView object (Excel)

Represents a view of a chart.


## Remarks

The  **ChartView** object is one of the objects that can be returned by the **[SheetViews](Excel.SheetViews.md)** collection, similar to the **[Sheets](Excel.Sheets.md)** collection. The **ChartView** object applies only to chart sheets.


## Example

The following example returns a  **ChartView** object.


```vb
ActiveWindow.SheetViews.Item(1) 

```

The following example returns a  **[Chart](Excel.Chart(object).md)** object.




```vb
ActiveWindow.SheetViews.Item(1).Sheet 

```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

