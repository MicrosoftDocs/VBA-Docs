---
title: WorksheetFunction.RTD method (Excel)
keywords: vbaxl10.chm137260
f1_keywords:
- vbaxl10.chm137260
ms.prod: excel
api_name:
- Excel.WorksheetFunction.RTD
ms.assetid: 1c3603d3-4f45-bd67-17f5-167685e3297c
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.RTD method (Excel)

This method connects to a source to receive real-time data (RTD).


## Syntax

_expression_.**RTD** (_progID_, _server_, _topic1_, _topic2_, _topic3_, _topic4_, _topic5_, _topic6_, _topic7_, _topic8_, _topic9_, _topic10_, _topic11_, _topic12_, _topic13_, _topic14_, _topic15_, _topic16_, _topic17_, _topic18_, _topic19_, _topic20_, _topic21_, _topic22_, _topic23_, _topic24_, _topic25_, _topic26_, _topic27_, _topic28_)

_expression_ An expression that returns a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _progID_|Required| **Variant**|A string representing the real-time server programmatic identifier.|
| _server_|Required| **Variant**|A server name, **Null** string, or **vbNullString** constant.|
| _topic1_|Required| **Variant**|A **String** representing a topic.|
| _topic2_ &ndash; _topic28_|Optional| **Variant**|A **String** representing a topic.|

## Return value

**Variant**


## Remarks

The _server_ argument is required in Visual Basic for Applications (VBA), even though it can be omitted on a worksheet.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]