---
title: Application.WorkbookPivotTableCloseConnection Event (Excel)
keywords: vbaxl10.chm504095
f1_keywords:
- vbaxl10.chm504095
ms.prod: excel
api_name:
- Excel.Application.WorkbookPivotTableCloseConnection
ms.assetid: 4c1d4cb2-f589-3c3c-ab4c-dcb08467fcfb
ms.date: 06/08/2017
---


# Application.WorkbookPivotTableCloseConnection Event (Excel)

Occurs after a PivotTable report connection has been closed.


## Syntax

 _expression_. `WorkbookPivotTableCloseConnection`( `_Wb_` , `_Target_` )

 _expression_ A variable that represents an [Application](Excel.Application(Graph property).md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](Excel.Workbook.md)**|The selected workbook.|
| _Target_|Required| **[PivotTable](Excel.PivotTable.md)**|The selected PivotTable report.|

### Return Value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been closed. This example assumes you have declared an object of type  **Workbook** with events in a class module.


```vb
Private Sub ConnectionApp_WorkbookPivotTableCloseConnection(ByVal wbOne As Workbook, Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```


## See also


[Application Object](Excel.Application(object).md)
[Workbook Object](Excel.Workbook.md)

