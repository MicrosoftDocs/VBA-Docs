---
title: Application.WorkbookOpen event (Excel)
keywords: vbaxl10.chm504081
f1_keywords:
- vbaxl10.chm504081
ms.prod: excel
api_name:
- Excel.Application.WorkbookOpen
ms.assetid: 37a5b55d-7968-29a2-3f87-edc3334c8ced
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.WorkbookOpen event (Excel)

Occurs when a workbook is opened.


## Syntax

_expression_.**WorkbookOpen** (_Wb_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The workbook.|

## Return value

Nothing


## Example

This example arranges all open windows when a workbook is opened.

```vb
Private Sub App_WorkbookOpen(ByVal Wb As Workbook) 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]