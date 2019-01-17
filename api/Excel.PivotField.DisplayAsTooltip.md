---
title: PivotField.DisplayAsTooltip property (Excel)
keywords: vbaxl10.chm240141
f1_keywords:
- vbaxl10.chm240141
ms.prod: excel
api_name:
- Excel.PivotField.DisplayAsTooltip
ms.assetid: 48e18a23-8e8c-828f-2522-71910dc54e2c
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotField.DisplayAsTooltip property (Excel)

This property is used to specify whether or not a specific member property PivotField is displayed in tooltips. Read/write  **Boolean**.


## Syntax

_expression_. `DisplayAsTooltip`

_expression_ A variable that represents a [PivotField](Excel.PivotField.md) object.


## Remarks

Trying to get or set these properties for PivotFields that are not member properties will return a run-time error.

A given member property is displayed in tooltips when the  **DisplayAsTooltip** property is set to **True** , and not displayed in tooltips when it is set to **False**. The default value is **True**.


## See also


[PivotField Object](Excel.PivotField.md)

