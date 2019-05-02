---
title: OLEObject.Verb method (Excel)
keywords: vbaxl10.chm417080
f1_keywords:
- vbaxl10.chm417080
ms.prod: excel
api_name:
- Excel.OLEObject.Verb
ms.assetid: c5714863-641c-1bfd-5688-9267494fb12d
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEObject.Verb method (Excel)

Sends a verb to the server of the specified OLE object.


## Syntax

_expression_.**Verb** (_Verb_)

_expression_ A variable that represents an **[OLEObject](Excel.OLEObject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Verb_|Optional| **[XlOLEVerb](Excel.XlOLEVerb.md)**|The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent.<br/><br/> The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the **XlOLEVerb** constants **xlOpen** and **xlPrimary**).|

## Return value

Variant


## Example

This example sends the default verb to the server for OLE object one on Sheet1.

```vb
Worksheets("Sheet1").OLEObjects(1).Verb
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]