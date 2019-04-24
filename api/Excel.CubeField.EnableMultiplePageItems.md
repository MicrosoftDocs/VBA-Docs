---
title: CubeField.EnableMultiplePageItems property (Excel)
keywords: vbaxl10.chm668090
f1_keywords:
- vbaxl10.chm668090
ms.prod: excel
api_name:
- Excel.CubeField.EnableMultiplePageItems
ms.assetid: 877328c6-dc30-e741-52ad-9cd91d7997c9
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.EnableMultiplePageItems property (Excel)

Set to **True** to allow multiple items in the page field area for OLAP PivotTables to be selected. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**EnableMultiplePageItems**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Example

This example determines if multiple page items are enabled for the cube field and notifies the user. The example assumes that an OLAP PivotTable exists on the active worksheet.

```vb
Sub UseMultiplePageItems() 
 
 Dim pvtTable As PivotTable 
 Dim cbeField As CubeField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set cbeField = pvtTable.CubeFields("[Country]") 
 
 ' Determine setting for multiple page items. 
 If cbeField.EnableMultiplePageItems = False Then 
 MsgBox "Multiple page items cannot be selected." 
 Else 
 MsgBox "Multiple page items can be selected." 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]