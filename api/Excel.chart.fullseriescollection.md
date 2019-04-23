---
title: Chart.FullSeriesCollection method (Excel)
keywords: vbaxl10.chm149194
f1_keywords:
- vbaxl10.chm149194
ms.prod: excel
ms.assetid: 875c18cf-064f-6b2f-2650-f5d07c16bc4d
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.FullSeriesCollection method (Excel)

Enables retrieving the filtered out series specified by the _Index_ argument.

## Syntax

_expression_.**FullSeriesCollection** (_Index_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant** |The indexed number of the filtered out **[Series](excel.series(object).md)** object.|

## Return value

**OBJECT**


## Remarks

**Series** objects in hidden rows or columns do not appear in the current series collection unless the user has enabled the **Show data in hidden rows and columns** option in the **Select Data** dialog.

> [!NOTE] 
> You can also use the series name in quotes:
>   
> _expression_.**FullSeriesCollection**(_"series name in quotes"_)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
