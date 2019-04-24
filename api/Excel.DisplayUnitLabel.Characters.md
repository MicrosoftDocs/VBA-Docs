---
title: DisplayUnitLabel.Characters property (Excel)
keywords: vbaxl10.chm673080
f1_keywords:
- vbaxl10.chm673080
ms.prod: excel
api_name:
- Excel.DisplayUnitLabel.Characters
ms.assetid: a136b4c9-be4a-9b17-20f6-c8b694202e9e
ms.date: 04/25/2019
localization_priority: Normal
---


# DisplayUnitLabel.Characters property (Excel)

Returns a **[Characters](Excel.Characters.md)** object that represents a range of characters within the object text. You can use the **Characters** object to format characters within a text string.


## Syntax

_expression_.**Characters** (_Start_, _Length_)

_expression_ A variable that represents a **[DisplayUnitLabel](excel.displayunitlabel(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the _Start_ character).|

## Remarks

The **Characters** object isn't a collection.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]