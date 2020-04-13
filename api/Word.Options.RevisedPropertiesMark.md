---
title: Options.RevisedPropertiesMark property (Word)
keywords: vbawd10.chm162988108
f1_keywords:
- vbawd10.chm162988108
ms.prod: word
api_name:
- Word.Options.RevisedPropertiesMark
ms.assetid: a973e504-3f27-a823-4746-d68b1b407413
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.RevisedPropertiesMark property (Word)

Returns or sets the mark used to show formatting changes while change tracking is enabled. Read/write  **WdRevisedPropertiesMark**.


## Syntax

_expression_. `RevisedPropertiesMark`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example causes text with changed formatting to be double-underlined when change tracking is enabled.


```vb
Options.RevisedPropertiesMark = _ 
 wdRevisedPropertiesMarkDoubleUnderline
```

This example returns the option selected in the **Formatting** box under **Track Changes** options on the **Track Changes** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.RevisedPropertiesMark
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]