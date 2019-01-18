---
title: Worksheet.EnableSelection property (Excel)
keywords: vbaxl10.chm175095
f1_keywords:
- vbaxl10.chm175095
ms.prod: excel
api_name:
- Excel.Worksheet.EnableSelection
ms.assetid: e1647c07-3863-9268-864c-1c62b2eebbb1
ms.date: 06/08/2017
localization_priority: Priority
---


# Worksheet.EnableSelection property (Excel)

Returns or sets what can be selected on the sheet. Read/write  **[xlEnableSelection](Excel.XlEnableSelection.md)**.


## Syntax

_expression_. `EnableSelection`

_expression_ A variable that represents a [Worksheet](./Excel.Worksheet.md) object.


## Remarks

This property takes effect only when the worksheet is protected:  **xlNoSelection** prevents any selection on the sheet, **xlUnlockedCells** allows only those cells whose **Locked** property is **False** to be selected, and **xlNoRestrictions** allows any cell to be selected.


## Example

This example sets worksheet one so that nothing on it can be selected.


```vb
With Worksheets(1) 
 .EnableSelection = xlNoSelection 
 .Protect Contents:=True, UserInterfaceOnly:=True 
End With
```


## See also


[Worksheet Object](Excel.Worksheet.md)

