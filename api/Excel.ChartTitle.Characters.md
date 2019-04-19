---
title: ChartTitle.Characters property (Excel)
keywords: vbaxl10.chm563081
f1_keywords:
- vbaxl10.chm563081
ms.prod: excel
api_name:
- Excel.ChartTitle.Characters
ms.assetid: 53cbfd6d-3be0-74d8-d55b-983d99d92e80
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartTitle.Characters property (Excel)

Returns a **[Characters](Excel.Characters.md)** object that represents a range of characters within the object text. You can use the **Characters** object to format characters within a text string.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

_expression_ A variable that represents a **[ChartTitle](Excel.ChartTitle(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the _Start_ character).|

## Remarks

The **Characters** object isn't a collection.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]