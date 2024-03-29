---
title: OLEObject.progID property (Excel)
keywords: vbaxl10.chm417083
f1_keywords:
- vbaxl10.chm417083
api_name:
- Excel.OLEObject.progID
ms.assetid: cbec1e95-6bdd-ce55-f426-28dcf4191897
ms.date: 05/02/2019
ms.localizationpriority: medium
---


# OLEObject.progID property (Excel)

Returns the programmatic identifiers for the object. Read-only **String**.


## Syntax

_expression_.**progID**

_expression_ A variable that represents an **[OLEObject](Excel.OLEObject.md)** object.


## Example

This example creates a list of the programmatic identifiers for the OLE objects on worksheet one.

```vb
rw = 0 
For Each o in Worksheets(1).OLEObjects 
 With Worksheets(2) 
 rw = rw + 1 
 .cells(rw, 1).Value = o.ProgId 
 End With 
Next
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]