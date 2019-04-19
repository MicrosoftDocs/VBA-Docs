---
title: ChartTitle.Text property (Excel)
keywords: vbaxl10.chm563087
f1_keywords:
- vbaxl10.chm563087
ms.prod: excel
api_name:
- Excel.ChartTitle.Text
ms.assetid: 22e073e3-06be-4888-cac3-7daad2a9cb33
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartTitle.Text property (Excel)

Returns or sets the text for the specified object. Read/write **String**.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a **[ChartTitle](Excel.ChartTitle(object).md)** object.


## Example

This example sets the text for the chart title of Chart1.


```vb
With Charts("Chart1") 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]