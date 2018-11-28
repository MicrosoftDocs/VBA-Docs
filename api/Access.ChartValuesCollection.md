---
title: ChartValuesCollection Object (Access)
keywords: vbaac10.chm14755
f1_keywords:
- vbaac10.chm14755
ms.prod: access
api_name:
- Access.ChartValuesCollection
ms.date: 05/02/2018
---


# ChartValuesCollection Object (Access)

A collection of all the **[ChartValues](Access.ChartValues.md)** objects in the specified chart.


## Using ChartValuesCollection

The following example displays the name of each **ChartValues** instance in a collection.

```vb
With myChart
 For Each cv In .ChartValuesCollection
  MsgBox (cv.Name)
 Next
End With
```

## See also

- [Chart object](Access.Chart.md)