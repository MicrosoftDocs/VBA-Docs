---
title: OLEFormat.Verb method (Excel)
keywords: vbaxl10.chm632076
f1_keywords:
- vbaxl10.chm632076
ms.prod: excel
api_name:
- Excel.OLEFormat.Verb
ms.assetid: bf5736e8-1909-ed0a-aaab-297ccde9ffef
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEFormat.Verb method (Excel)

Sends a verb to the server of the specified OLE object.


## Syntax

_expression_.**Verb** (_Verb_)

_expression_ A variable that represents an **[OLEFormat](Excel.OLEFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Verb_|Optional| **[XlOLEVerb](Excel.XlOLEVerb.md)**|The verb that the server of the OLE object should act on. If this argument is omitted, the default verb is sent.<br/><br/>The available verbs are determined by the object's source application. Typical verbs for an OLE object are Open and Primary (represented by the **XlOLEVerb** constants **xlOpen** and **xlPrimary**).|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]