---
title: ChartObject.PrintObject property (Excel)
keywords: vbaxl10.chm494089
f1_keywords:
- vbaxl10.chm494089
ms.prod: excel
api_name:
- Excel.ChartObject.PrintObject
ms.assetid: 504f4a82-6129-cb38-ea2f-f9b29e14d036
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.PrintObject property (Excel)

**True** if the object will be printed when the document is printed. Read/write **Boolean**.


## Syntax

_expression_.**PrintObject**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example sets embedded chart one on Sheet1 to be printed with the worksheet.

```vb
Worksheets("Sheet1").ChartObjects(1).PrintObject = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]