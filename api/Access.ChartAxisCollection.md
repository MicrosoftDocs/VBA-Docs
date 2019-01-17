---
title: ChartAxisCollection object (Access)
keywords: vbaac10.chm14753
f1_keywords:
- vbaac10.chm14753
ms.prod: access
api_name:
- Access.ChartAxisCollection
ms.date: 11/28/2018
localization_priority: Normal
---


# ChartAxisCollection object (Access)

A collection of all the **[ChartAxis](Access.ChartAxis.md)** objects in the specified chart.


## Example

The following example displays the number of axes in the collection, and then displays the name of each axis.

```vb
With myChart
 MsgBox (.ChartAxisCollection.Count)
  For Each axis In .ChartAxisCollection
    MsgBox (axis.Name)
  Next
End With
```

## See also

- [Chart object](Access.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]