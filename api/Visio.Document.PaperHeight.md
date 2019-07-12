---
title: Document.PaperHeight property (Visio)
keywords: vis_sdr.chm10514015
f1_keywords:
- vis_sdr.chm10514015
ms.prod: visio
api_name:
- Visio.Document.PaperHeight
ms.assetid: 305356e8-69d6-bae3-5136-d931fcf967b5
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.PaperHeight property (Visio)

Returns the height of a document's printed page. Read-only.


## Syntax

_expression_.**PaperHeight** (_UnitsNameOrCode_)

_expression_ A variable that represents a **[Document](Visio.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UnitsNameOrCode_|Required| **Variant**|The units to use when retrieving the paper height.|

## Return value

Double


## Remarks

The **PaperHeight** property value can be a string such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Visio units such as centimeters, meters, miles, and so on. You can also use any of the unit constants declared by the Visio type library in **[VisUnitCodes](Visio.visunitcodes.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]