---
title: Workbook.PivotTableCloseConnection Event (Excel)
keywords: vbaxl10.chm503094
f1_keywords:
- vbaxl10.chm503094
ms.prod: excel
ms.assetid: e267ab5b-382e-b270-18c8-f643e03e4604
ms.date: 06/08/2017
localization_priority: Normal
---


# Workbook.PivotTableCloseConnection Event (Excel)

Occurs after a PivotTable report closes the connection to its data source.


## Syntax

_expression_. `PivotTableCloseConnection`( `_Target_` )

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[PivotTable](Excel.PivotTable.md)**|The selected PivotTable report.|
| _Target_|Required|PIVOTTABLE||
|Name|Required/Optional|Data type|Description|

## Return value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been closed. This example assumes you have declared an object of type  **Workbook** with events in a class module.


```vb
Private Sub ConnectionApp_PivotTableCloseConnection(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been closed." 
 
End Sub
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]