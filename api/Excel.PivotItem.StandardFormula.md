---
title: PivotItem.StandardFormula property (Excel)
keywords: vbaxl10.chm246092
f1_keywords:
- vbaxl10.chm246092
ms.prod: excel
api_name:
- Excel.PivotItem.StandardFormula
ms.assetid: 34410ff5-0330-f685-e508-94084e6f0e5d
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotItem.StandardFormula property (Excel)

Returns or sets a  **String** specifying formulas with standard English (United States) formatting. Read/write.


## Syntax

_expression_. `StandardFormula`

_expression_ A variable that represents a [PivotItem](Excel.PivotItem.md) object.


## Remarks

The  **StandardFormula** property primarily affects item names with date or number formatting. It provides a way to specify or query a formula for a given calculated item.

The  **[StandardFormula](Excel.PivotItem.StandardFormula.md)** property is "international-friendly" whereas the **[Formula](Excel.PivotItem.Formula.md)** property is not.


## Example

This example adds 10 to the Decimals field and displays it as a calculated item in the data field. The example assumes that a PivotTable exists on the active worksheet and that a field titled "Decimals" exists in the data table.


```vb
Sub UseStandardFomula() 
 
 Dim pvtTable As PivotTable 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Change calculated field of decimals by adding '10'. 
 pvtTable.CalculatedFields.Item(1).StandardFormula = "Decimals + 10" 
 
End Sub
```


## See also


[PivotItem Object](Excel.PivotItem.md)

