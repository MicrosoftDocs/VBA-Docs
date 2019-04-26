---
title: InsideTop property (Excel Graph)
keywords: vbagr10.chm67204
f1_keywords:
- vbagr10.chm67204
ms.prod: excel
api_name:
- Excel.InsideTop
ms.assetid: c91d9788-0eb4-02ed-48f0-2118d317b1ec
ms.date: 04/11/2019
localization_priority: Normal
---


# InsideTop property (Excel Graph)

Returns the distance from the chart edge to the inside top edge of the plot area, in [points](../language/glossary/vbe-glossary.md#point). Read-only **Double**.

## Syntax

_expression_.**InsideTop**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

The plot area used for this measurement doesn't include the axis labels. The **[Top](excel.top.md)** property for the plot area uses the bounding rectangle that includes the axis labels.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]