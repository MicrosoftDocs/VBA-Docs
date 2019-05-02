---
title: OLEDBError.Stage property (Excel)
keywords: vbaxl10.chm654077
f1_keywords:
- vbaxl10.chm654077
ms.prod: excel
api_name:
- Excel.OLEDBError.Stage
ms.assetid: 71dd6495-3d03-307d-a7cd-816779f25754
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBError.Stage property (Excel)

Returns a numeric value specifying the stage of an error that resulted after the most recent OLE DB query. Read-only **Long**.


## Syntax

_expression_.**Stage**

_expression_ A variable that represents an **[OLEDBError](Excel.OLEDBError.md)** object.


## Example

This example displays the error numbers, stage, and other error information returned by the most recent OLE DB query.

```vb
Set objEr = Application.OLEDBErrors(1) 
MsgBox "The following error occurred:" & _ 
 objEr.Number & ", " & objEr.Native & ", " & _ 
 objEr.Stage & ", " & _ 
 objEr.ErrorString & " : " & objEr.SqlState
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]