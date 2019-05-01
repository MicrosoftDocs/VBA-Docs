---
title: OLEObject.Update method (Excel)
keywords: vbaxl10.chm417079
f1_keywords:
- vbaxl10.chm417079
ms.prod: excel
api_name:
- Excel.OLEObject.Update
ms.assetid: 7784b688-fef2-14b3-761a-df412dfa0282
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEObject.Update method (Excel)

Updates the link.


## Syntax

_expression_.**Update**

_expression_ A variable that represents an **[OLEObject](Excel.OLEObject.md)** object.


## Return value

Variant


## Example

This example updates the link to OLE object one on Sheet1.

```vb
Worksheets("Sheet1").OLEObjects(1).Update
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]