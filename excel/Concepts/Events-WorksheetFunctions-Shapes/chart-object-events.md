---
title: Chart object events
keywords: vbaxl10.chm5199755
f1_keywords:
- vbaxl10.chm5199755
ms.prod: excel
ms.assetid: 6808dfde-94d0-afb0-b245-44d8d1d6241e
ms.date: 11/13/2018
localization_priority: Normal
---


# Chart object events

Chart events occur when the user activates or changes a chart. Events on chart sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and select **View Code** from the shortcut menu. Select the event name from the **Procedure** drop-down list box.

- [Activate](../../../api/Excel.Chart.Activate(even).md) 
- [BeforeDoubleClick](../../../api/Excel.Chart.BeforeDoubleClick.md) 
- [BeforeRightClick](../../../api/Excel.Chart.BeforeRightClick.md) 
- [Calculate](../../../api/Excel.Chart.Calculate.md) 
- [Deactivate](../../../api/Excel.Chart.Deactivate.md) 
- [MouseDown](../../../api/Excel.Chart.MouseDown.md) 
- [MouseMove](../../../api/Excel.Chart.MouseMove.md) 
- [MouseUp](../../../api/Excel.Chart.MouseUp.md) 
- [Resize](../../../api/Excel.Chart.Resize.md) 
- [Select](../../../api/Excel.Chart.Select(even).md) 
- [SeriesChange](../../../api/Excel.Chart.SeriesChange.md)

> [!NOTE] 
> To write event procedures for an embedded chart, you must create a new object using the **WithEvents** keyword in a class module. For more information, see [Using events with embedded charts](using-events-with-embedded-charts.md).

This example changes a point's border color when the user changes the point value.

```vb
Private Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _ 
        ByVal PointIndex As Long) 
    Set p = ActiveChart.SeriesCollection(SeriesIndex). _ 
        Points(PointIndex) 
    p.Border.ColorIndex = 3 
End Sub
```

## See also

- [Excel functions (by category)](https://support.office.com/article/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]