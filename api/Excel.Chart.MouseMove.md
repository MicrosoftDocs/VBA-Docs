---
title: Chart.MouseMove event (Excel)
keywords: vbaxl10.chm500078
f1_keywords:
- vbaxl10.chm500078
ms.prod: excel
api_name:
- Excel.Chart.MouseMove
ms.assetid: b1277953-a882-f00f-2ac1-dd0cc49fef72
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.MouseMove event (Excel)

Occurs when the position of the mouse pointer changes over a chart.


## Syntax

_expression_.**MouseMove** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Long**|The mouse button that was released. Can be one of the following  **[XlMouseButton](Excel.XlMouseButton.md)** constants: **xlNoButton**, **xlPrimaryButton**, or **xlSecondaryButton**.|
| _Shift_|Required| **Long**|The state of the SHIFT, CTRL, and ALT keys when the event occurred. Can be one of or a sum of values.|
| _x_|Required| **Long**|The X coordinate of the mouse pointer in chart object client coordinates.|
| _y_|Required| **Long**|The Y coordinate of the mouse pointer in chart object client coordinates.|

## Return value

Nothing


## Remarks

The following table specifies the values for the  _Shift_ parameter.



|Value|Description|
|:-----|:-----|
|0 (zero)|No keys|
|1|SHIFT key|
|2|CTRL key|
|4|ALT key|

## Example

This example runs when the position of the mouse pointer changes over a chart.


```vb
Private Sub Chart_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 MsgBox "X = " & X & " Y = " & Y 
End Sub
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]