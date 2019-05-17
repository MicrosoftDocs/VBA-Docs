---
title: Style.IncludePatterns property (Excel)
keywords: vbaxl10.chm177084
f1_keywords:
- vbaxl10.chm177084
ms.prod: excel
api_name:
- Excel.Style.IncludePatterns
ms.assetid: edb7e87f-20d2-2bea-b2e8-83ffab749e3e
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.IncludePatterns property (Excel)

**True** if the style includes the **Color**, **ColorIndex**, **InvertIfNegative**, **Pattern**, **PatternColor**, and **PatternColorIndex** properties of the **[Interior](excel.interior(object).md)** object. Read/write **Boolean**.


## Syntax

_expression_.**IncludePatterns**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example sets the style attached to cell A1 on Sheet1 to include pattern format.

```vb
Worksheets("Sheet1").Range("A1").Style.IncludePatterns = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]