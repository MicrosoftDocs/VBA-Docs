---
title: PivotTable.ColumnFields property (Excel)
keywords: vbaxl10.chm235074
f1_keywords:
- vbaxl10.chm235074
ms.prod: excel
api_name:
- Excel.PivotTable.ColumnFields
ms.assetid: caae2016-e213-31f0-5ce7-fd8593ad4266
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ColumnFields property (Excel)

Returns an object that represents either a single PivotTable field (a **[PivotField](Excel.PivotField.md)** object) or a collection of all the fields (a **[PivotFields](Excel.PivotFields.md)** object) that are currently shown as column fields. Read-only.


## Syntax

_expression_.**ColumnFields** (_Index_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The field name or number (can be an array to specify more than one field).|

## Example

This example adds the field names of the PivotTable report columns to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.ColumnFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
