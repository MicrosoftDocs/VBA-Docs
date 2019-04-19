---
title: Chart.SetElement method (Excel)
keywords: vbaxl10.chm149175
f1_keywords:
- vbaxl10.chm149175
ms.prod: excel
api_name:
- Excel.Chart.SetElement
ms.assetid: 0efff437-179b-fe16-118b-6f3cde49c5cf
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.SetElement method (Excel)

Sets chart elements on a chart. Read/write **[MsoChartElementType](office.msochartelementtype.md)**.


## Syntax

_expression_.**SetElement** (_Element_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Element_|Required| **MsoChartElementType**|Specifies the chart element type.|

## Return value

Nothing


## Remarks

For charts, the following commands in the **Layout** tab correspond to the **SetElement** method:

- Everything in the **Labels** group.
    
- Everything in the **Axes** group.
    
- Everything in the **Analysis** group.
    
- **PlotArea**, **Chart Wall**, and **Chart Floor** buttons.
    

**MsoChartElementType** is an enumeration of constants that refer to all of the above commands.


## Example

This example sets chart elements by using the various constant values to an active chart.

```vb
ActiveChart.Axes(xlValue).MajorGridlines.Select 
 ActiveChart.SetElement (msoElementChartTitleCenteredOverlay) 
 ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinor) 
 ActiveChart.Walls.Select 
 Application.CommandBars("Clip Art").Visible = False 
 ActiveChart.SetElement (msoElementChartFloorShow)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]