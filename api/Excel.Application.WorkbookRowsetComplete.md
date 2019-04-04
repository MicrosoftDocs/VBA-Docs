---
title: Application.WorkbookRowsetComplete event (Excel)
keywords: vbaxl10.chm504102
f1_keywords:
- vbaxl10.chm504102
ms.prod: excel
api_name:
- Excel.Application.WorkbookRowsetComplete
ms.assetid: cc472400-5622-5b4f-60a2-d3347ded266f
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookRowsetComplete event (Excel)

The **WorkbookRowsetComplete** event occurs when the user either drills through the recordset or invokes the rowset action on an OLAP PivotTable.


## Syntax

_expression_.**WorkbookRowsetComplete** (_Wb_, _Description_, _Sheet_, _Success_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook for which the event occurs.|
| _Description_|Required| **String**|A brief description of the event.|
| _Sheet_|Required| **String**|The worksheet on which the recordset is created.|
| _Success_|Required| **Boolean**|Contains a **Boolean** value to indicate success or failure.|

## Remarks

Because the recordset is created asynchronously, the event allows automation to determine when the action has been completed. Additionally, because the recordset is created on a separate sheet, the event needs to be on the workbook level.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]