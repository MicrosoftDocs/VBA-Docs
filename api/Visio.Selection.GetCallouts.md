---
title: Selection.GetCallouts method (Visio)
keywords: vis_sdr.chm11162170
f1_keywords:
- vis_sdr.chm11162170
ms.prod: visio
api_name:
- Visio.Selection.GetCallouts
ms.assetid: 29adcbbc-d5a9-a284-c025-785ad1ccf2c8
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.GetCallouts method (Visio)

Returns the list of identifiers of the callout shapes in the selection.


## Syntax

_expression_. `GetCallouts`( `_NestedOptions_` )

_expression_ A variable that represents a **[Selection](Visio.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NestedOptions_|Required| **[VisContainerNested](Visio.VisContainerNested.md)**|Indicates whether to exclude shapes in the selection that are contained by containers or lists. See Remarks for possible values.|

## Return value

 **Long()**


## Remarks

The  _NestedOptions_ parameter must be one of the following **VisContainerNested** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visContainerIncludeNested**|0|Include shapes that are in nested containers.|
| **visContainerExcludeNested**|1|Exclude shapes that are in nested containers.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]