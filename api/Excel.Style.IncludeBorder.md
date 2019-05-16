---
title: Style.IncludeBorder property (Excel)
keywords: vbaxl10.chm177081
f1_keywords:
- vbaxl10.chm177081
ms.prod: excel
api_name:
- Excel.Style.IncludeBorder
ms.assetid: 81b44216-e8fa-88fe-e82c-7fd8844d33ea
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.IncludeBorder property (Excel)

**True** if the style includes the **Color**, **ColorIndex**, **LineStyle**, and **Weight** properties of the **[Border](excel.border(object).md)** object. Read/write **Boolean**.


## Syntax

_expression_.**IncludeBorder**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include border format.

```vb
Worksheets("Sheet1").Range("A1").Style.IncludeBorder = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]