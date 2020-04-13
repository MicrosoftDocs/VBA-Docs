---
title: ChartBorder.LineStyle property (Word)
keywords: vbawd10.chm61014020
f1_keywords:
- vbawd10.chm61014020
ms.prod: word
api_name:
- Word.ChartBorder.LineStyle
ms.assetid: f11e0877-2a3c-4aa6-471f-333d6b485249
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartBorder.LineStyle property (Word)

Returns or sets the line style for the border. Read/write  **[XlLineStyle](Word.xllinestyle.md)**, **xlGray25**, **xlGray50**, **xlGray75**, or **xlAutomatic**.


## Syntax

_expression_.**LineStyle**

_expression_ A variable that represents a '[ChartBorder](Word.ChartBorder.md)' object.


## Remarks

The **xlDouble** and **xlSlantDashDot** constants of the **xlLineStyle** enumeration do not apply to charts.


## Example

The following example puts a border around the chart area and the plot area of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .ChartArea.Border.LineStyle = xlDashDot 
 With .PlotArea.Border 
 .LineStyle = xlDashDotDot 
 .Weight = xlThick 
 End With 
 End With 
 End If 
End With 

```


## See also


[ChartBorder Object](Word.ChartBorder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]