---
title: InsideLeft property (Excel Graph)
keywords: vbagr10.chm5207555
f1_keywords:
- vbagr10.chm5207555
ms.prod: excel
api_name:
- Excel.InsideLeft
ms.assetid: 04c9291b-efbf-4deb-d6b4-373473531ba6
ms.date: 04/11/2019
localization_priority: Normal
---


# InsideLeft property (Excel Graph)

Returns the distance from the chart edge to the inside left edge of the plot area, in [points](../language/glossary/vbe-glossary.md#point). Read-only **Double**.

## Syntax

_expression_.**InsideLeft**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

The plot area used for this measurement doesn't include the axis labels. The **[Left](excel.left.md)** property for the plot area uses the bounding rectangle that includes the axis labels.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]