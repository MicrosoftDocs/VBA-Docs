---
title: PivotCache.RefreshPeriod property (Excel)
keywords: vbaxl10.chm227091
f1_keywords:
- vbaxl10.chm227091
ms.prod: excel
api_name:
- Excel.PivotCache.RefreshPeriod
ms.assetid: 6357769c-e73e-2388-962a-f3bb790c423e
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.RefreshPeriod property (Excel)

Returns or sets the number of minutes between refreshes. Read/write **Long**.


## Syntax

_expression_.**RefreshPeriod**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

Setting the period to 0 (zero) disables automatic timed refreshes and is equivalent to setting this property to **Null**.

The value of the **RefreshPeriod** property can be an integer from 0 through 32767.


## Example

This example sets the refresh period for the PivotTable cache (PivotTable3) to 15 minutes.

```vb
Set objPC = Worksheets("Sheet1").PivotTables("PivotTable3").PivotCache 
objPC.RefreshPeriod = 15
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]