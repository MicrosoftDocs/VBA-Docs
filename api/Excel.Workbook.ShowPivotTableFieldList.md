---
title: Workbook.ShowPivotTableFieldList property (Excel)
keywords: vbaxl10.chm199196
f1_keywords:
- vbaxl10.chm199196
ms.prod: excel
api_name:
- Excel.Workbook.ShowPivotTableFieldList
ms.assetid: 33c74c54-27ea-d230-c640-47109bdfd4a2
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ShowPivotTableFieldList property (Excel)

**True** (default) if the PivotTable field list can be shown. Read/write **Boolean**.


## Syntax

_expression_.**ShowPivotTableFieldList**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

In this example, Microsoft Excel determines if the PivotTable field list can be shown and notifies the user.

```vb
Sub UseShowPivotTableFieldList() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.ActiveWorkbook 
 
 'Determine PivotTable field list setting. 
 If wkbOne.ShowPivotTableFieldList = True Then 
 MsgBox "The PivotTable field list can be viewed." 
 Else 
 MsgBox "The PivotTable field list cannot be viewed." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]