---
title: Document.TopMargin property (Visio)
keywords: vis_sdr.chm10514580
f1_keywords:
- vis_sdr.chm10514580
ms.prod: visio
api_name:
- Visio.Document.TopMargin
ms.assetid: ed8d16c2-f80d-d444-28a4-d9f0db4ab6d3
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.TopMargin property (Visio)

Specifies the top margin when printing a document. Read/write.


## Syntax

_expression_.**TopMargin** (_UnitsNameOrCode_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Optional| **Variant**|The units to use when retrieving or setting the margin value.|

## Return value

Double


## Remarks

If  _UnitsNameOrCode_ is not provided, the **TopMargin** property will default to internal drawing units (inches).

The **TopMargin** property corresponds to the **Top** setting in the **Print Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow, and then click **Setup** on the **Print Setup** tab).

Units can be an integer or string value such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Visio units such as centimeters, meters, miles, and so on. You can also use any of the units constants declared by the Visio type library in **[VisUnitCodes](Visio.visunitcodes.md)**.

For a list of valid integer and string values, see [About units of measure](../visio/Concepts/about-units-of-measure-visio.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]