---
title: Application.BoxAlign method (Project)
keywords: vbapj.chm29
f1_keywords:
- vbapj.chm29
ms.prod: project-server
api_name:
- Project.Application.BoxAlign
ms.assetid: 2b27c9a0-36fa-1bbd-96e3-267b95ad5407
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BoxAlign method (Project)

Aligns the specified part of the selected boxes in the active Network Diagram view with the same part of the box that has the focus.


## Syntax

_expression_. `BoxAlign`( `_Alignment_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Alignment_|Required|**Long**|Specifies which side or portion of a box to use for the alignment. Can be one of the  **[PjAlign](Project.PjAlign.md)** constants.|

## Return value

 **Boolean**


## Remarks

If only one box is selected, the  **BoxAlign** method has no effect.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]