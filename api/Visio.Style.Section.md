---
title: Style.Section Property (Visio)
keywords: vis_sdr.chm11414300
f1_keywords:
- vis_sdr.chm11414300
ms.prod: visio
api_name:
- Visio.Style.Section
ms.assetid: 932acfc4-9713-4c7c-0472-a160ebddeecc
ms.date: 06/08/2017
---


# Style.Section Property (Visio)

Returns the requested  **Section** object belonging to a style. Read-only.


## Syntax

 _expression_. `Section`( `_Index_` )

 _expression_ A variable that represents a [Style](./Visio.Style.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**|A section index.|

## Return value

Section


## Remarks

Constants that represent sections are prefixed with  **visSection** and are declared by the Microsoft Visio type library in **[VisSectionIndices](Visio.vissectionindices.md)**.


