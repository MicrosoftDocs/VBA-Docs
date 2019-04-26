---
title: InsideHeight property (Excel Graph)
keywords: vbagr10.chm5207553
f1_keywords:
- vbagr10.chm5207553
ms.prod: excel
api_name:
- Excel.InsideHeight
ms.assetid: 02528324-3aaf-17b3-984d-96ab7b446d5a
ms.date: 04/11/2019
localization_priority: Normal
---


# InsideHeight property (Excel Graph)

Returns the inside height of the plot area, in [points](../language/glossary/vbe-glossary.md#point). Read-only **Double**.

## Syntax

_expression_.**InsideHeight**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

The plot area used for this measurement doesn't include the axis labels. The **[Height](excel.height.md)** property for the plot area uses the bounding rectangle that includes the axis labels.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]