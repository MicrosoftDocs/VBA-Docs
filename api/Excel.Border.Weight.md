---
title: Border.Weight property (Excel)
keywords: vbaxl10.chm547076
f1_keywords:
- vbaxl10.chm547076
ms.prod: excel
api_name:
- Excel.Border.Weight
ms.assetid: c6b9a812-60e6-245d-e86e-fb385581f890
ms.date: 03/07/2019
localization_priority: Normal
---


# Border.Weight property (Excel)

Returns or sets an **[XlBorderWeight](Excel.XlBorderWeight.md)** value that represents the weight of the border.


## Syntax

_expression_.**Weight**

_expression_ A variable that represents a **[Border](Excel.Border(object).md)** object.

## Remarks

> [!IMPORTANT] 
> Note that the visual properties of a **Border** object are interlocked; that is, changing one property can induce changes in another. In most cases, the induced changes serve to make the border visible (which may or may not be desirable). However, other (more unexpected) results are possible. For an example, see the **[Border](excel.border(object).md)** object.

## Example

This example sets the border weight for oval one on Sheet1.

```vb
Worksheets("Sheet1").Ovals(1).Border.Weight = xlMedium
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
