---
title: PivotTable.DataFields property (Excel)
keywords: vbaxl10.chm235079
f1_keywords:
- vbaxl10.chm235079
ms.prod: excel
api_name:
- Excel.PivotTable.DataFields
ms.assetid: 32f9f635-c247-ad1b-6bb8-6eef4f03dc67
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.DataFields property (Excel)

Returns an object that represents either a single PivotTable field (a **[PivotField](Excel.PivotField.md)** object) or a collection of all the fields (a **[PivotFields](Excel.PivotFields.md)** object) that are currently shown as data fields. Read-only.


## Syntax

_expression_.**DataFields** (_Index_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The field name or number (can be an array to specify more than one field).|

## Example

This example adds the names for the PivotTable data fields to a list on a new worksheet.

```vb
Set nwSheet = Worksheets.Add 
nwSheet.Activate 
Set pvtTable = Worksheets("Sheet2").Range("A1").PivotTable 
rw = 0 
For Each pvtField In pvtTable.DataFields 
 rw = rw + 1 
 nwSheet.Cells(rw, 1).Value = pvtField.Name 
Next pvtField
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
