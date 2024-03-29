---
title: PageSetup.Order property (Excel)
keywords: vbaxl10.chm473089
f1_keywords:
- vbaxl10.chm473089
api_name:
- Excel.PageSetup.Order
ms.assetid: cb39c83a-3ab2-cab9-531c-762db4ab6770
ms.date: 05/03/2019
ms.localizationpriority: medium
---


# PageSetup.Order property (Excel)

Returns or sets an **[XlOrder](Excel.XlOrder.md)** value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.


## Syntax

_expression_.**Order**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example breaks Sheet1 into pages when the worksheet is printed. Numbering and printing proceed from the first page to the pages to the right, and then move down and continue printing across the sheet.

```vb
Worksheets("Sheet1").PageSetup.Order = xlOverThenDown
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]