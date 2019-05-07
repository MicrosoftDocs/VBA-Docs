---
title: Shape.SectionExists property (Visio)
keywords: vis_sdr.chm11214305
f1_keywords:
- vis_sdr.chm11214305
ms.prod: visio
api_name:
- Visio.Shape.SectionExists
ms.assetid: 588a3b17-4831-b7bb-455f-12bc5c3620fc
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.SectionExists property (Visio)

Determines whether a ShapeSheet section exists for a particular shape. Read-only.


## Syntax

_expression_. `SectionExists`( `_Section_` , `_fExistsLocally_` )

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Section_|Required| **Integer**|The section index.|
| _fExistsLocally_|Required| **Integer**|The scope of the search.|

## Return value

Integer


## Remarks

If  _fExistsLocally_ is **False** (0), the **SectionExists** property returns **True** if the object either contains or inherits the section. If _fExistsLocally_ is **True** (non-zero), the **SectionExists** property returns **True** only if the object contains the section locally; if the section is inherited, the **SectionExists** property returns **False**.

Constants that represent sections are prefixed with  **visSection** and are declared by the Microsoft Visio type library in **[VisSectionIndices](Visio.vissectionindices.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]