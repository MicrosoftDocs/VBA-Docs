---
title: Sequences.Add method (PowerPoint)
keywords: vbapp10.chm650004
f1_keywords:
- vbapp10.chm650004
ms.prod: powerpoint
api_name:
- PowerPoint.Sequences.Add
ms.assetid: 5f1516ec-d617-ffcf-c786-318a7ba3cb1e
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequences.Add method (PowerPoint)

Returns a  **[Sequence](PowerPoint.Sequence.md)** object that represents a new sequence.


## Syntax

_expression_.**Add** (_Index_)

_expression_ A variable that represents a [Sequences](PowerPoint.Sequences.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Long**|The position of the sequence in relation to other sequences. The default value is -1, which means that if you omit the Index parameter, the new sequence is added to the end of the existing sequences.|

## Return value

Sequence


## See also


[Sequences Object](PowerPoint.Sequences.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]