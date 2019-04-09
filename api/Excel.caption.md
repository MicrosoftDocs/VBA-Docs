---
title: Caption property (Excel Graph)
keywords: vbagr10.chm3076968
f1_keywords:
- vbagr10.chm3076968
ms.prod: excel
ms.assetid: 37d9afab-873c-c026-fb76-33987aa103b8
ms.date: 04/10/2019
localization_priority: Normal
---


# Caption property (Excel Graph)

Returns or sets the title text for the object. Read/write **String**.

## Syntax

_expression_.**Caption**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example adds the title Annual Salary Figures to the chart.

```vb
myChart.HasTitle = True 
myChart.ChartTitle.Caption = "Annual Salary Figures" 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]