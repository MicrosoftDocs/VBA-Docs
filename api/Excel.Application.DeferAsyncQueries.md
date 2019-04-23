---
title: Application.DeferAsyncQueries property (Excel)
keywords: vbaxl10.chm133313
f1_keywords:
- vbaxl10.chm133313
ms.prod: excel
api_name:
- Excel.Application.DeferAsyncQueries
ms.assetid: 21f05a5a-40e8-304a-f537-41ea171a114c
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DeferAsyncQueries property (Excel)

Gets or sets whether asynchronous queries to OLAP data sources are executed when a worksheet is calculated by VBA code. Read/write **Boolean**.


## Syntax

_expression_.**DeferAsyncQueries**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

Setting the **DeferAsyncQueries** property to **True** prevents any asynchronous queries from executing until the **[CalculateUntilAsyncQueriesDone](Excel.Application.CalculateUntilAsyncQueriesDone.md)** method is called.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]