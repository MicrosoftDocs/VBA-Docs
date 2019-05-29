---
title: Workbook.Close method (Excel)
keywords: vbaxl10.chm199085
f1_keywords:
- vbaxl10.chm199085
ms.prod: excel
api_name:
- Excel.Workbook.Close
ms.assetid: c0376cab-a2db-c606-67bf-0a4921b81e03
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Close method (Excel)

Closes the object.


## Syntax

_expression_.**Close** (_SaveChanges_, _FileName_, _RouteWorkbook_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveChanges_|Optional| **Variant**|If there are no changes to the workbook, this argument is ignored. If there are changes to the workbook and the workbook appears in other open windows, this argument is ignored. If there are changes to the workbook but the workbook doesn't appear in any other open windows, this argument specifies whether changes should be saved. If set to **True**, changes are saved to the workbook.<br/><br/>If there is not yet a file name associated with the workbook, _FileName_ is used. If _FileName_ is omitted, the user is asked to supply a file name.|
| _FileName_|Optional| **Variant**|Saves changes under this file name.|
| _RouteWorkbook_|Optional| **Variant**|If the workbook doesn't need to be routed to the next recipient (if it has no routing slip or has already been routed), this argument is ignored. Otherwise, Microsoft Excel routes the workbook according to the value of this parameter.<br/><br/>If set to **True**, the workbook is sent to the next recipient. If set to **False**, the workbook is not sent. If omitted, the user is asked whether the workbook should be sent.|

## Remarks

Closing a workbook from Visual Basic doesn't run any Auto_Close macros in the workbook. Use the **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run the Auto_Close macros.


## Example

This example closes Book1.xls and discards any changes that have been made to it.

```vb
Workbooks("BOOK1.XLS").Close SaveChanges:=False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
