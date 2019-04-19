---
title: ColorStops.Count property (Excel)
keywords: vbaxl10.chm853073
f1_keywords:
- vbaxl10.chm853073
ms.prod: excel
api_name:
- Excel.ColorStops.Count
ms.assetid: 0574a698-ff87-56e3-eea9-aa2e6e77f270
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorStops.Count property (Excel)

Returns or sets the count of the represented object. Read-only.


## Syntax

_expression_.**Count**

_expression_ An expression that returns a **[ColorStops](Excel.ColorStops.md)** object.


## Return value

Long


## Example

Returns the number of **ColorStops** in the active cell.

```vb
ActiveCell.Interior.Gradient.ColorStops.Count
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]