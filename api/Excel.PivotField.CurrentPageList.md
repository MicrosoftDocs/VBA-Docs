---
title: PivotField.CurrentPageList property (Excel)
keywords: vbaxl10.chm240135
f1_keywords:
- vbaxl10.chm240135
api_name:
- Excel.PivotField.CurrentPageList
ms.assetid: 3efde5a2-4cf3-b95d-e7ba-65ea8e184e64
ms.date: 05/04/2019
ms.localizationpriority: medium
---


# PivotField.CurrentPageList property (Excel)

Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. Read/write **Variant**.


## Syntax

_expression_.**CurrentPageList**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

To avoid run-time errors, the data source must be an OLAP source, the field chosen must currently be in the Page position, and the **[EnableMultiplePageItems](Excel.PivotField.EnableMultiplePageItems.md)** property must be set to **True**.


## Example

This example sets the page field to list the Food items of the PivotTable report. It assumes that a PivotTable exists on the active worksheet.

```vb
Sub UseCurrentPageList() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields("[Product]") 
 
 ' To avoid run-time errors set the following property to True. 
 pvtTable.CubeFields("[Product]").EnableMultiplePageItems = True 
 
 ' Set the page list to "Food". 
 pvtField.CurrentPageList = "[Product].[All Products].[Food]" 
 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]