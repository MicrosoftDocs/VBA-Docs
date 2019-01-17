---
title: AutoCorrect.DeleteReplacement method (Excel)
keywords: vbaxl10.chm545075
f1_keywords:
- vbaxl10.chm545075
ms.prod: excel
api_name:
- Excel.AutoCorrect.DeleteReplacement
ms.assetid: 765e207d-64b3-c85d-ae10-937eaf836e0a
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect.DeleteReplacement method (Excel)

Deletes an entry from the array of AutoCorrect replacements.


## Syntax

_expression_. `DeleteReplacement`( `_What_` )

_expression_ A variable that represents an [AutoCorrect](Excel.AutoCorrect-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _What_|Required| **String**|The text to be replaced, as it appears in the row to be deleted from the array of AutoCorrect replacements. If this string doesn't exist in the array of AutoCorrect replacements, this method fails.|

## Return value

Variant


## Example

This example removes the word "Temperature" from the array of AutoCorrect replacements.


```vb
With Application.AutoCorrect 
 .DeleteReplacement "Temperature" 
End With
```


## See also


[AutoCorrect Object](Excel.AutoCorrect(object).md)

