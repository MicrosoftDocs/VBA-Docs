---
title: Chart.MouseUp event (Excel)
keywords: vbaxl10.chm500077
f1_keywords:
- vbaxl10.chm500077
api_name:
- Excel.Chart.MouseUp
ms.assetid: 45281aac-a4f6-390d-e767-a4fe2ee670fc
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.MouseUp event (Excel)

Occurs when a mouse button is released while the pointer is over a chart.


## Syntax

_expression_.**MouseUp** (_Button_, _Shift_, _x_, _y_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Button_|Required| **Long**|The mouse button that was released. Can be one of the following **[XlMouseButton](Excel.XlMouseButton.md)** constants: **xlNoButton**, **xlPrimaryButton**, or **xlSecondaryButton**.|
| _Shift_|Required| **Long**|The state of the Shift, Ctrl, and Alt keys when the event occurred. Can be one of or a sum of values.|
| _x_|Required| **Long**|The _x_ coordinate of the mouse pointer in chart object client coordinates.|
| _y_|Required| **Long**|The _y_ coordinate of the mouse pointer in chart object client coordinates.|

## Return value

Nothing


## Remarks

The following table specifies the values for the _Shift_ parameter.

|Value|Description|
|:-----|:-----|
|0 (zero)|No keys|
|1|Shift key|
|2|Ctrl key|
|4|Alt key|

## Example

This example runs when a mouse button is released over a chart.

```vb
Private Sub Chart_MouseUp(ByVal Button As Long, _ 
 ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 MsgBox "Button = " & Button & chr$(13) & _ 
 "Shift = " & Shift & chr$(13) & _ 
 "X = " & X & " Y = " & Y 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]