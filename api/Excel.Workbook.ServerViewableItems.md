---
title: Workbook.ServerViewableItems property (Excel)
keywords: vbaxl10.chm199245
f1_keywords:
- vbaxl10.chm199245
ms.prod: excel
api_name:
- Excel.Workbook.ServerViewableItems
ms.assetid: 2c10a647-2b2c-0640-9990-109b89444cd2
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ServerViewableItems property (Excel)

Allows a developer to interact with the list of published objects in the workbook that are shown on the server. Read-only.


## Syntax

_expression_.**ServerViewableItems**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

This property returns a collection of the items that could be published to Excel Services. It can include **Tables**, **PivotTables**, **Named Ranges**, and **Charts**. It can also contain **Sheets** as long as it is not a mixture of **Sheets** and the before mentioned list.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]