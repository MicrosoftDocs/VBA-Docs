---
title: PivotCache.RobustConnect property (Excel)
keywords: vbaxl10.chm227105
f1_keywords:
- vbaxl10.chm227105
ms.prod: excel
api_name:
- Excel.PivotCache.RobustConnect
ms.assetid: 354d0124-e178-342b-9565-fa74e9dae5d5
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.RobustConnect property (Excel)

Returns or sets how the PivotTable cache connects to its data source. Read/write **[XlRobustConnect](Excel.XlRobustConnect.md)**.


## Syntax

_expression_.**RobustConnect**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Example

The following example determines the setting for the cache connection and notifies the user. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub CheckRobustConnect() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the connection robustness and notify user. 
 Select Case pvtCache.RobustConnect 
 Case xlAlways 
 MsgBox "The PivotTable cache is always connected to its source." 
 Case xlAsRequired 
 MsgBox "The PivotTable cache is connected to its source as required." 
 Case xlNever 
 MsgBox "The PivotTable cache is never connected to its source." 
 End Select 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]