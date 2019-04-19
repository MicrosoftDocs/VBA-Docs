---
title: Chart.SeriesChange event (Excel)
keywords: vbaxl10.chm500084
f1_keywords:
- vbaxl10.chm500084
ms.prod: excel
api_name:
- Excel.Chart.SeriesChange
ms.assetid: 80a8058c-0445-0051-24d1-1a965c302790
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.SeriesChange event (Excel)

Occurs when the user changes the value of a chart data point by choosing a bar in the chart and dragging the top edge up or down thus changing the value of the data point.

> [!IMPORTANT] 
> This event is not functional in Excel 2007 and later versions. You should not use it in your code.


## Syntax

_expression_.**SeriesChange** (_SeriesIndex_, _PointIndex_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SeriesIndex_|Required| **Long**| The offset within the **[Series](Excel.Series(object).md)** collection for the changed series.|
| _PointIndex_|Required| **Long**|The offset within the **[Points](Excel.Points(object).md)** collection for the changed point.|

## Return value

Nothing


## Example

This example changes the point's border color when the user changes the point value.

```vb
Private Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _ 
 ByVal PointIndex As Long) 
 Set p = Me.SeriesCollection(SeriesIndex).Points(PointIndex) 
 p.Border.ColorIndex = 3 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]