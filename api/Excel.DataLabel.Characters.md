---
title: DataLabel.Characters property (Excel)
keywords: vbaxl10.chm582081
f1_keywords:
- vbaxl10.chm582081
ms.prod: excel
api_name:
- Excel.DataLabel.Characters
ms.assetid: 0072e034-727d-6de5-f2bc-ce398ac750bc
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.Characters property (Excel)

Returns a  **[Characters](Excel.Characters.md)** object that represents a range of characters within the object text. You can use the **Characters** object to format characters within a text string.


## Syntax

_expression_. `Characters`( `_Start_` , `_Length_` )

_expression_ A variable that represents a [DataLabel](Excel.DataLabel-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first character to be returned. If this argument is either 1 or omitted, this property returns a range of characters starting with the first character.|
| _Length_|Optional| **Variant**|The number of characters to be returned. If this argument is omitted, this property returns the remainder of the string (everything after the  _Start_ character).|

## Remarks

The  **Characters** object isn't a collection.


## See also


[DataLabel Object](Excel.DataLabel(object).md)

