---
title: PivotCache.CommandText property (Excel)
keywords: vbaxl10.chm227087
f1_keywords:
- vbaxl10.chm227087
ms.prod: excel
api_name:
- Excel.PivotCache.CommandText
ms.assetid: 07921bda-74fe-2a41-15f7-16068ce49a31
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.CommandText property (Excel)

Returns or sets the command string for the specified data source. Read/write **Variant**.


## Syntax

_expression_.**CommandText**

_expression_ An expression that returns a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

For OLE DB sources, the **[CommandType](Excel.PivotCache.CommandType.md)** property describes the value of the **CommandText** property.

For ODBC sources, setting the **CommandText** property causes the data to be refreshed.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]