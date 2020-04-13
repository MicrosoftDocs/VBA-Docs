---
title: Application.PresentationBeforeClose event (PowerPoint)
keywords: vbapp10.chm621025
f1_keywords:
- vbapp10.chm621025
ms.prod: powerpoint
api_name:
- PowerPoint.Application.PresentationBeforeClose
ms.assetid: 8c2d820b-aa44-287b-10ad-1dc6f4122231
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PresentationBeforeClose event (PowerPoint)

Represents a **Presentation** object before it closes.


## Syntax

_expression_. `PresentationBeforeClose`( `_Pres_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The **Presentation** object.|
| _Cancel_|Required|**Boolean**|If set to  **True**, the presentation will not close.|

## Return value

Nothing


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]