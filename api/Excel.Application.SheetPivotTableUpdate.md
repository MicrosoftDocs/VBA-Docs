---
title: Application.SheetPivotTableUpdate event (Excel)
keywords: vbaxl10.chm504094
f1_keywords:
- vbaxl10.chm504094
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableUpdate
ms.assetid: f42d1f7b-6395-326b-4b4f-72b497c81bd3
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.SheetPivotTableUpdate event (Excel)

Occurs after the sheet of the PivotTable report has been updated.


## Syntax

_expression_.**SheetPivotTableUpdate** (_Sh_, _Target_)

_expression_ An expression that returns an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The selected sheet.|
| _Target_|Required| **[PivotTable](Excel.PivotTable.md)**|The selected PivotTable report.|

## Example

This example displays a message stating that the sheet of the PivotTable report has been updated. This example assumes that you have declared an object of type **Application** or **[Workbook](Excel.Workbook.md)** with events in a class module.

```vb
Private Sub ConnectionApp_SheetPivotTableUpdate(ByVal shOne As Object, Target As PivotTable) 
 
 MsgBox "The SheetPivotTable connection has been updated." 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]