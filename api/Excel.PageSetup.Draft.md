---
title: PageSetup.Draft property (Excel)
keywords: vbaxl10.chm473080
f1_keywords:
- vbaxl10.chm473080
ms.prod: excel
api_name:
- Excel.PageSetup.Draft
ms.assetid: 133d474c-2058-7dd9-d10b-0e45d9b2f972
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.Draft property (Excel)

 **True** if the sheet will be printed without graphics. Read/write **Boolean**.


## Syntax

_expression_. `Draft`

_expression_ A variable that represents a [PageSetup](Excel.PageSetup.md) object.


## Remarks

Setting this property to  **True** makes printing faster (at the expense of not printing graphics).


## Example

This example turns off graphics printing for Sheet1.


```vb
Worksheets("Sheet1").PageSetup.Draft = True
```


## See also


[PageSetup Object](Excel.PageSetup.md)

