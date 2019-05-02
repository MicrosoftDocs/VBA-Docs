---
title: PageSetup.FirstPageNumber property (Excel)
keywords: vbaxl10.chm473081
f1_keywords:
- vbaxl10.chm473081
ms.prod: excel
api_name:
- Excel.PageSetup.FirstPageNumber
ms.assetid: 606d2bb3-9e3f-2d98-01ea-3257e83f61ea
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.FirstPageNumber property (Excel)

Returns or sets the first page number that will be used when this sheet is printed. If **xlAutomatic**, Microsoft Excel chooses the first page number. The default is **xlAutomatic** ([Constants](excel.constants.md)). Read/write **Long**.


## Syntax

_expression_.**FirstPageNumber**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example sets the first page number of Sheet1 to 100.

```vb
Worksheets("Sheet1").PageSetup.FirstPageNumber = 100
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]