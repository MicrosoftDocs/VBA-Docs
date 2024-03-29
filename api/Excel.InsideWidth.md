---
title: InsideWidth property (Excel Graph)
keywords: vbagr10.chm5207559
f1_keywords:
- vbagr10.chm5207559
api_name:
- Excel.InsideWidth
ms.assetid: 1f6bfd65-c134-6d52-5936-dfc4a4eecda8
ms.date: 04/11/2019
ms.localizationpriority: medium
---


# InsideWidth property (Excel Graph)

Returns the inside width of the plot area, in [points](../language/glossary/vbe-glossary.md#point). Read-only **Double**.

## Syntax

_expression_.**InsideWidth**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

The plot area used for this measurement doesn't include the axis labels. The **[Width](excel.width.md)** property for the plot area uses the bounding rectangle that includes the axis labels.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]