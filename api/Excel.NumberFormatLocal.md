---
title: NumberFormatLocal property (Excel Graph)
keywords: vbagr10.chm66633
f1_keywords:
- vbagr10.chm66633
ms.prod: excel
api_name:
- Excel.NumberFormatLocal
ms.assetid: 186aee6a-dd66-39a7-cebc-546c3e156d6d
ms.date: 04/12/2019
localization_priority: Normal
---


# NumberFormatLocal property (Excel Graph)

Returns or sets the format code for the specified object as a string in the language of the user. Read/write **Variant**.

## Syntax

_expression_.**NumberFormatLocal**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

The **Format** function uses different format code strings than do the **[NumberFormat](excel.numberformat.md)** and **NumberFormatLocal** properties.


## Example

This example displays the number format for the data labels in the first series on the chart in the language of the user.

```vb
MsgBox "The number format for the first series is " & _ 
 myChart.SeriesCollection(1).DataLabels.NumberFormatLocal
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]