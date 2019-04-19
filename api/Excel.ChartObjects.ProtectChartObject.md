---
title: ChartObjects.ProtectChartObject property (Excel)
keywords: vbaxl10.chm497098
f1_keywords:
- vbaxl10.chm497098
ms.prod: excel
api_name:
- Excel.ChartObjects.ProtectChartObject
ms.assetid: e0685fbd-84a5-36c4-a5ab-06127937f2c8
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObjects.ProtectChartObject property (Excel)

**True** if the embedded chart frame cannot be moved, resized, or deleted through the user interface. Read/write **Boolean**.


## Syntax

_expression_.**ProtectChartObject**

_expression_ A variable that represents a **[ChartObjects](Excel.ChartObjects.md)** object.


## Remarks

Setting this property to **True** will not protect the embedded chart frame from being modified through the object model.


## Example

```vb
Worksheets(1).ChartObjects(1).ProtectChartObject = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]