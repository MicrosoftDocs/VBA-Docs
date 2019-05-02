---
title: PageSetup.PrintGridlines property (Excel)
keywords: vbaxl10.chm473093
f1_keywords:
- vbaxl10.chm473093
ms.prod: excel
api_name:
- Excel.PageSetup.PrintGridlines
ms.assetid: 9a92c157-046a-00b5-3813-a5c924ac42b9
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.PrintGridlines property (Excel)

**True** if cell gridlines are printed on the page. Applies only to worksheets. Read/write **Boolean**.


## Syntax

_expression_.**PrintGridlines**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example prints cell gridlines when Sheet1 is printed.

```vb
Worksheets("Sheet1").PageSetup.PrintGridlines = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]