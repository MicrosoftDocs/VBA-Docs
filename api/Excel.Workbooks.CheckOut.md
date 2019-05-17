---
title: Workbooks.CheckOut method (Excel)
keywords: vbaxl10.chm203085
f1_keywords:
- vbaxl10.chm203085
ms.prod: excel
api_name:
- Excel.Workbooks.CheckOut
ms.assetid: 11b9eb2a-8c9a-6e61-63e2-554030243388
ms.date: 05/18/2019
localization_priority: Normal
---


# Workbooks.CheckOut method (Excel)

Returns a **String** representing a specified workbook from a server to a local computer for editing.


## Syntax

_expression_.**CheckOut** (_FileName_)

_expression_ A variable that represents a **[Workbooks](Excel.Workbooks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to check out.|

## Example

This example verifies that a workbook is not checked out by another user and can be checked out. If the workbook can be checked out, it copies the workbook to the local computer for editing.

```vb
Sub UseCheckOut(docCheckOut As String) 
 
 ' Determine if workbook can be checked out. 
 If Workbooks.CanCheckOut(docCheckOut) = True Then 
 Workbooks.CheckOut docCheckOut 
 Else 
 MsgBox "Unable to check out this document at this time." 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
