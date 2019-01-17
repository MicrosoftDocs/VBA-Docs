---
title: Application.ShowChartTipValues property (Excel)
keywords: vbaxl10.chm133209
f1_keywords:
- vbaxl10.chm133209
ms.prod: excel
api_name:
- Excel.Application.ShowChartTipValues
ms.assetid: 886b2cf9-f6b3-3770-3082-28f2f99863cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowChartTipValues property (Excel)

 **True** if charts show chart tip values. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_. `ShowChartTipValues`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example turns off chart tip names and values.


```vb
With Application 
 .ShowChartTipNames = False 
 .ShowChartTipValues = False 
End With
```


## See also


[Application Object](Excel.Application(object).md)

