---
title: PivotCell.DataField property (Excel)
keywords: vbaxl10.chm692075
f1_keywords:
- vbaxl10.chm692075
ms.prod: excel
api_name:
- Excel.PivotCell.DataField
ms.assetid: d5373236-ba29-301a-2c49-ccda89c69328
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotCell.DataField property (Excel)

Returns a  **[PivotField](Excel.PivotField.md)** object that corresponds to the selected data field.


## Syntax

_expression_. `DataField`

_expression_ A variable that represents a [PivotCell](Excel.PivotCell.md) object.


## Remarks

This property will return an error if the  **PivotCell** object is not one of the allowed constants of **[xlPivotCellType](Excel.XlPivotCellType.md)** : **xlPivotCellTypeDataField** , **xlPivotCellTypeSubtotal** , or **xlPivotCellTypeGrandTotal**.


## Example

This example determines if cell L10 is in the data field of the PivotTable and either returns the PivotTable field that corresponds to the cell by notifying the user, or handles the run-time error. The example assumes a PivotTable exists in the active worksheet.


```vb
Sub CheckDataField() 
 
 On Error GoTo Not_In_DataField 
 
 MsgBox Application.Range("L10").PivotCell.DataField 
 Exit Sub 
 
Not_In_DataField: 
 MsgBox "The selected range is not in the data field of the PivotTable." 
 
End Sub
```


## See also


[PivotCell Object](Excel.PivotCell.md)

