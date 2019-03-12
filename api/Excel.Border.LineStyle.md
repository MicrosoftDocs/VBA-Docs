---
title: Border.LineStyle property (Excel)
keywords: vbaxl10.chm547075
f1_keywords:
- vbaxl10.chm547075
ms.prod: excel
api_name:
- Excel.Border.LineStyle
ms.assetid: 7f2529b7-4782-8d8d-d529-6d8d19417db4
ms.date: 03/07/2019
localization_priority: Normal
---


# Border.LineStyle property (Excel)

Returns or sets the line style for the border. Read/write **[XlLineStyle](Excel.XlLineStyle.md)**, **xlGray25**, **xlGray50**, **xlGray75**, or **xlAutomatic**.


## Syntax

_expression_.**LineStyle**

_expression_ A variable that represents a **[Border](Excel.Border(object).md)** object.


## Remarks

**xlDouble** and **xlSlantDashDot** do not apply to charts.

> [!IMPORTANT] 
> Note that the visual properties of a **Border** object are interlocked; that is, changing one property can induce changes in another. In most cases, the induced changes serve to make the border visible (which may or may not be desirable). However, other (more unexpected) results are possible. For an example, see the **[Border](excel.border(object).md)** object.

## Example

This example puts a border around the chart area and the plot area of Chart1.

```vb
With Charts("Chart1") 
 .ChartArea.Border.LineStyle = xlDashDot 
 With .PlotArea.Border 
 .LineStyle = xlDashDotDot 
 .Weight = xlThick 
 End With 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
