---
title: Application.DeferAsyncQueries property (Excel)
keywords: vbaxl10.chm133313
f1_keywords:
- vbaxl10.chm133313
ms.prod: excel
api_name:
- Excel.Application.DeferAsyncQueries
ms.assetid: 21f05a5a-40e8-304a-f537-41ea171a114c
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DeferAsyncQueries property (Excel)

Gets or sets whether asynchronous queries to OLAP data sources are executed when a worksheet is calculated by VBA code. Read/write  **Boolean**.


## Syntax

_expression_. `DeferAsyncQueries`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

Setting the  **DeferAsyncQueries** property to **True** prevents any asynchronous queries form executing until the **[CalculateUntilAsyncQueriesDone](Excel.Application.CalculateUntilAsyncQueriesDone.md)** method is called.


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]