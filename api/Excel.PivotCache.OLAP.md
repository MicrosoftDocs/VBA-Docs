---
title: PivotCache.OLAP property (Excel)
keywords: vbaxl10.chm227100
f1_keywords:
- vbaxl10.chm227100
ms.prod: excel
api_name:
- Excel.PivotCache.OLAP
ms.assetid: d40d3a71-0a27-c4a6-0c3b-47ab7a1a0e06
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.OLAP property (Excel)

Returns **True** if the PivotTable cache is connected to an Online Analytical Processing (OLAP) server. Read-only **Boolean**.


## Syntax

_expression_.**OLAP**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Example

This example determines whether the cache connection is to an OLAP server. The example assumes that a PivotTable exists on the active worksheet.

```vb
Sub CheckPivotCache() 
 
 ' Determine if PivotCache has OLAP connection. 
 If Application.ActiveWorkbook.PivotCaches.Item(1).OLAP = True Then 
 MsgBox "The PivotCache is connected to an OLAP server" 
 Else 
 MsgBox "The PivotCache is not connected to an OLAP server." 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]