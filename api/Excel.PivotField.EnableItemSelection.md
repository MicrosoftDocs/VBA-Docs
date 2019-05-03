---
title: PivotField.EnableItemSelection property (Excel)
keywords: vbaxl10.chm240134
f1_keywords:
- vbaxl10.chm240134
ms.prod: excel
api_name:
- Excel.PivotField.EnableItemSelection
ms.assetid: ae55f88a-618f-3063-2019-a993a3146b67
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.EnableItemSelection property (Excel)

When set to **False**, disables the ability to use the field dropdown in the user interface. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**EnableItemSelection**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

A run-time error occurs if the OLAP PivotTable field is not the highest level for the hierarchy.


## Example

This example determines the setting for selecting items by using the field dropdown and enables the feature, if necessary. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub UseEnableItemSelection() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.RowFields(1) 
 
 ' Determine setting for property and enable if necessary. 
 If pvtField.EnableItemSelection = False Then 
 pvtField.EnableItemSelection = True 
 MsgBox "Item selection enabled for fields." 
 Else 
 MsgBox "Item selection is already enabled for fields." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]