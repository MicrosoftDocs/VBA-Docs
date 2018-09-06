---
title: PivotLayout Object (Excel)
keywords: vbaxl10.chm663072
f1_keywords:
- vbaxl10.chm663072
ms.prod: excel
api_name:
- Excel.PivotLayout
ms.assetid: cfef617e-f49a-e969-7873-40593412a32e
ms.date: 06/08/2017
---


# PivotLayout Object (Excel)

Represents the placement of fields in a PivotChart report.


## Example

Use the  **[PivotLayout](Excel.Chart.PivotLayout.md)** property to return a **PivotLayout** object. The following example creates a list of PivotTable field names used in the first PivotChart report.


```vb
Sub ListFieldNames 
 
 Dim objNewSheet As Worksheet 
 Dim intRow As Integer 
 Dim objPF As PivotField 
 
 Set objNewSheet = Worksheets.Add 
 
 intRow = 1 
 
 For Each objPF In _ 
 Charts("Chart1").PivotLayout.PivotFields 
 
 objNewSheet.Cells(intRow, 1).Value = objPF.Caption 
 
 intRow = intRow + 1 
 
 Next objPF 
 
End Sub
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)


