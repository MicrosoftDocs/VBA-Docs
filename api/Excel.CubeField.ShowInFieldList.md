---
title: CubeField.ShowInFieldList property (Excel)
keywords: vbaxl10.chm668092
f1_keywords:
- vbaxl10.chm668092
ms.prod: excel
api_name:
- Excel.CubeField.ShowInFieldList
ms.assetid: 9a9163f3-b398-5059-9dce-b993413e850b
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.ShowInFieldList property (Excel)

When set to **True** (default), a **CubeField** object will be shown in the field list. Read/write **Boolean**.


## Syntax

_expression_.**ShowInFieldList**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Example

In this example, Microsoft Excel determines if a **CubeField** object can be shown in the Field list and notifies the user. This example assumes that a PivotTable report exists on the active worksheet and that a **CubeField** object exists.

```vb
Sub IsCubeFieldInList() 
 
 Dim pvtTable As PivotTable 
 Dim cbeField As CubeField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set cbeField = pvtTable.CubeFields("[Country]") 
 
 ' Determine if a CubeField can be seen. 
 If cbeField.ShowInFieldList = True Then 
 MsgBox "The CubeField object can be seen in the field list." 
 Else 
 MsgBox "The CubeField object cannot be seen in the field list." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]