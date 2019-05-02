---
title: PageSetup.CenterVertically property (Excel)
keywords: vbaxl10.chm473078
f1_keywords:
- vbaxl10.chm473078
ms.prod: excel
api_name:
- Excel.PageSetup.CenterVertically
ms.assetid: ffd5897b-fe25-f52f-eb94-cb42659bcedd
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.CenterVertically property (Excel)

**True** if the sheet is centered vertically on the page when it's printed. Read/write **Boolean**.


## Syntax

_expression_.**CenterVertically**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example centers Sheet1 vertically when it's printed.

```vb
Worksheets("Sheet1").PageSetup.CenterVertically = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]