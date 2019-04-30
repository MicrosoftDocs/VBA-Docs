---
title: PivotTable.DataPivotField property (Excel)
keywords: vbaxl10.chm235140
f1_keywords:
- vbaxl10.chm235140
ms.prod: excel
api_name:
- Excel.PivotTable.DataPivotField
ms.assetid: 00b62ffd-76bd-cd4b-218c-b6d695150efb
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotTable.DataPivotField property (Excel)

Returns a  **[PivotField](Excel.PivotField.md)** object that represents all the data fields in a PivotTable. Read-only.


## Syntax

_expression_. `DataPivotField`

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Example

This example moves the second  **PivotItem** object to the first position. It assumes a PivotTable exists on the active worksheet and that the PivotTable contains data fields.


```vb
Sub UseDataPivotField() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Move second PivotItem to the first position in PivotTable. 
 pvtTable.DataPivotField.PivotItems(2).Position = 1 
 
End Sub
```


## See also


[PivotTable Object](Excel.PivotTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]