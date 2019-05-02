---
title: PageSetup.CenterHorizontally property (Excel)
keywords: vbaxl10.chm473077
f1_keywords:
- vbaxl10.chm473077
ms.prod: excel
api_name:
- Excel.PageSetup.CenterHorizontally
ms.assetid: 6b3e97fd-6b05-6863-c642-b085ea9ddd33
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.CenterHorizontally property (Excel)

**True** if the sheet is centered horizontally on the page when it's printed. Read/write **Boolean**.


## Syntax

_expression_.**CenterHorizontally**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example centers Sheet1 horizontally when it's printed.

```vb
Worksheets("Sheet1").PageSetup.CenterHorizontally = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
