---
title: OLEDBConnection.BackgroundQuery property (Excel)
keywords: vbaxl10.chm794074
f1_keywords:
- vbaxl10.chm794074
ms.prod: excel
api_name:
- Excel.OLEDBConnection.BackgroundQuery
ms.assetid: c106c0d8-16ea-83dc-1b4e-1a311d9c0d9e
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.BackgroundQuery property (Excel)

**True** if queries for the OLE DB connection are performed asynchronously (in the background). Read/write **Boolean**.


## Syntax

_expression_.**BackgroundQuery**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Remarks

For OLAP data sources, this property is read-only and always returns **False**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]