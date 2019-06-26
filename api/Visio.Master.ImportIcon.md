---
title: Master.ImportIcon method (Visio)
keywords: vis_sdr.chm10716360
f1_keywords:
- vis_sdr.chm10716360
ms.prod: visio
api_name:
- Visio.Master.ImportIcon
ms.assetid: 886d724d-9d02-ab6f-8049-80fa04f8caec
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.ImportIcon method (Visio)

Imports the icon for a  **Master** object from a named file.


## Syntax

_expression_. `ImportIcon`( `_FileName_` )

_expression_ A variable that represents a **[Master](Visio.Master.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to import.|

## Return value

Nothing


## Remarks

The  **ImportIcon** method can only import files that were produced by exporting a master icon in the application's internal icon format (**visIconFormatVisio**); it does not accept icons in other file formats.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]