---
title: ODBCErrors.Item method (Excel)
keywords: vbaxl10.chm529074
f1_keywords:
- vbaxl10.chm529074
ms.prod: excel
api_name:
- Excel.ODBCErrors.Item
ms.assetid: 694a0e7e-f6c0-8721-792b-8e82e6a8e5c1
ms.date: 05/01/2019
localization_priority: Normal
---


# ODBCErrors.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[ODBCErrors](Excel.ODBCErrors.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number for the object.|

## Return value

An  **[ODBCError](Excel.ODBCError.md)** object contained by the collection.


## Example

This example displays an ODBC error.


```vb
Set er = Application.ODBCErrors.Item(1) 
MsgBox "The following error occurred:" & 
 er.ErrorString & " : " & er.SqlState
```


## See also


[ODBCErrors Object](Excel.ODBCErrors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]