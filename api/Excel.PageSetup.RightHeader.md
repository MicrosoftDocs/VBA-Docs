---
title: PageSetup.RightHeader property (Excel)
keywords: vbaxl10.chm473100
f1_keywords:
- vbaxl10.chm473100
ms.prod: excel
api_name:
- Excel.PageSetup.RightHeader
ms.assetid: 97e1780d-d511-d433-0e31-501381e6318d
ms.date: 05/03/2019
localization_priority: Normal
---


# PageSetup.RightHeader property (Excel)

Returns or sets the right part of the header. Read/write **String**.


## Syntax

_expression_.**RightHeader**

_expression_ A variable that represents a **[PageSetup](Excel.PageSetup.md)** object.


## Example

This example prints the file name in the upper-right corner of every page.

```vb
Worksheets("Sheet1").PageSetup.RightHeader = "&F"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]