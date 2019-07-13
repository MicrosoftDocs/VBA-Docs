---
title: Document.HeaderMargin property (Visio)
keywords: vis_sdr.chm10550650
f1_keywords:
- vis_sdr.chm10550650
ms.prod: visio
api_name:
- Visio.Document.HeaderMargin
ms.assetid: 7d2c137d-6b75-9747-5a6a-5e5d99156d45
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.HeaderMargin property (Visio)

Gets or sets the margin of a document's header. Read/write.


## Syntax

_expression_.**HeaderMargin** (_UnitsNameOrCode_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|The units to use when retrieving or setting the cell's value. Defaults to internal drawing units (inches).|

## Return value

Double


## Remarks

You can also set this value in the **Margin** box under **Header** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click **Print Preview**, and then in the **Preview** group, click **Header & Footer**).

Automation constants for representing units are declared by the Visio type library in member **[VisUnitCodes](Visio.visunitcodes.md)**.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]