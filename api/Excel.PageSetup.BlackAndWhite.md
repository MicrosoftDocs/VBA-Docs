---
title: PageSetup.BlackAndWhite property (Excel)
keywords: vbaxl10.chm473073
f1_keywords:
- vbaxl10.chm473073
ms.prod: excel
api_name:
- Excel.PageSetup.BlackAndWhite
ms.assetid: 81d1fd09-d317-7d3f-5200-875340a5917e
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.BlackAndWhite property (Excel)

**True** if elements of the document will be printed in black and white. Read/write **Boolean**.


## Syntax

_expression_.**BlackAndWhite**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Remarks

This property applies only to worksheet pages.


## Example

This example causes Sheet1 to be printed in black and white.

```vb
Worksheets("Sheet1").PageSetup.BlackAndWhite = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]