---
title: Designs.Add method (PowerPoint)
keywords: vbapp10.chm643004
f1_keywords:
- vbapp10.chm643004
ms.prod: powerpoint
api_name:
- PowerPoint.Designs.Add
ms.assetid: 00608390-a12b-d698-36a6-ded2df3cc26a
ms.date: 06/08/2017
localization_priority: Normal
---


# Designs.Add method (PowerPoint)

Returns a **[Design](PowerPoint.Design.md)** object that represents a new slide design.


## Syntax

_expression_.**Add** (_designName_, _Index_)

_expression_ A variable that represents a [Designs](PowerPoint.Designs.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _designName_|Required|**String**|The name of the design.|
| _Index_|Optional|**Integer**|The index number of the design in the  **Designs** collection. The default value is -1, which means that if you omit the Index parameter, the new slide design is added at the end of existing slide designs.|

## Return value

Design


## See also


[Designs Object](PowerPoint.Designs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]