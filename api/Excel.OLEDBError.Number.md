---
title: OLEDBError.Number property (Excel)
keywords: vbaxl10.chm654076
f1_keywords:
- vbaxl10.chm654076
ms.prod: excel
api_name:
- Excel.OLEDBError.Number
ms.assetid: 9e88a0bb-1cbf-d98e-52a9-a8f9a0bde81c
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBError.Number property (Excel)

Returns a numeric value that specifies an error. The error number corresponds to a unique trap number corresponding to an error condition that resulted after the most recent OLE DB query. Read-only **Long**.


## Syntax

_expression_.**Number**

_expression_ A variable that represents an **[OLEDBError](Excel.OLEDBError.md)** object.


## Example

This example displays the error number and other error information returned by the most recent OLE DB query.

```vb
Set objEr = Application.OLEDBErrors(1) 
MsgBox "The following error occurred:" & _ 
 objEr.Number & ", " & objEr.Native & ", " & _ 
 objEr.ErrorString & " : " & objEr.SqlState
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]