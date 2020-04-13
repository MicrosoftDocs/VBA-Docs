---
title: Chart.HasAxis property (Project)
ms.prod: project-server
ms.assetid: f1059a7e-01ac-cd41-78d6-dc88f52943f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.HasAxis property (Project)
Gets or sets which axes exist on a chart. Read/write  **Variant**.

## Syntax

_expression_.**HasAxis**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _axisType_|Required|**Variant**|The axis type. Series axes apply only to 3D charts. Can be one of the **Office.XlAxisType** constants.|
| _AxisGroup_|Optional|**Variant**|The axis group. 3D charts have only one set of axes. Can be one of the **Office.XlAxisGroup** constants.|

## Return value

 **Period**


## Remarks

You must enter a value for at least one of the parameters when setting the **HasAxis** property.

Project may create or delete axes if you change the chart type or the **IMsoAxis.AxisGroup**,  **IMsoChartGroup.AxisGroup**, or  **IMsoSeries.AxisGroup** properties.


## Example

The following example turns on the primary value axis for the chart.


```vb
Sub SetPrimaryAxis()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.HasAxis(Office.XlAxisType.xlValue, Office.XlAxisType.xlPrimary) = True
End Sub
```


## Property value

 **VARIANT**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]