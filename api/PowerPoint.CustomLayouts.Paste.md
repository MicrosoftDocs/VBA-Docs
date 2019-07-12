---
title: CustomLayouts.Paste method (PowerPoint)
keywords: vbapp10.chm671005
f1_keywords:
- vbapp10.chm671005
ms.prod: powerpoint
api_name:
- PowerPoint.CustomLayouts.Paste
ms.assetid: d4fcd2db-3d6b-0c59-6ea3-f9aadf90ed04
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomLayouts.Paste method (PowerPoint)

Pastes the slides on the Clipboard into a custom layout and adds the custom layout to the  **[CustomLayouts](PowerPoint.CustomLayouts.md)** collection.


## Syntax

_expression_.**Paste** (_Index_)

_expression_ A variable that represents a [CustomLayouts](PowerPoint.CustomLayouts.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Long**|The index number of the custom layout before which the new custom layout is pasted. If this argument is omitted, the new custom layout is pasted at the end of the  **CustomLayouts** collection.|

## Return value

CustomLayout


## See also


[CustomLayouts Object](PowerPoint.CustomLayouts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]