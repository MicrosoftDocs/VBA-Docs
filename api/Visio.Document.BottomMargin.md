---
title: Document.BottomMargin property (Visio)
keywords: vis_sdr.chm10513150
f1_keywords:
- vis_sdr.chm10513150
ms.prod: visio
api_name:
- Visio.Document.BottomMargin
ms.assetid: 5fd185a5-ecc9-000e-f5b0-fa309d52847a
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.BottomMargin property (Visio)

Specifies the bottom margin when printing the pages in a document. Read/write.


## Syntax

_expression_.**BottomMargin** (_UnitsNameOrCode_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|he units to use when retrieving or setting the margin value. Defaults to internal drawing units.|

## Return value

Double


## Remarks

The value of this property corresponds to the value entered in the **Bottom** box in the **Print Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow, and then click **Setup** on the **Print Setup** tab).

You can specify  _UnitsNameOrCode_ as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all set _UnitsNameOrCode_ to inches.

- **ActiveDocument.BottomMargin** (**visInches**) = _newValue_

- **ActiveDocument.BottomMargin** (65) = _newValue_

- **ActiveDocument.BottomMargin** ("in") = _newValue_ where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "i".

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

Automation constants for representing units are declared by the Microsoft Visio type library in member **[VisUnitCodes](Visio.visunitcodes.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]