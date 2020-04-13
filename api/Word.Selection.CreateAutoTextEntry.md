---
title: Selection.CreateAutoTextEntry method (Word)
keywords: vbawd10.chm158663190
f1_keywords:
- vbawd10.chm158663190
ms.prod: word
api_name:
- Word.Selection.CreateAutoTextEntry
ms.assetid: def6f758-af70-eaf2-f15c-4a6a28c247b5
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.CreateAutoTextEntry method (Word)

Adds a new  **[AutoTextEntry](Word.AutoTextEntry.md)** object to the **[AutoTextEntries](Word.autotextentries.md)** collection, based on the current selection.


## Syntax

_expression_. `CreateAutoTextEntry`( `_Name_` , `_StyleName_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The text the user must type to call the new AutoText entry.|
| _StyleName_|Required| **String**|The category in which the new AutoText entry will be listed on the **AutoText** menu.|

## Example

This example creates a new AutoText entry named "handdel" under the category "Mailing Instructions," given "HAND DELIVERY" as the currently selected text.


```vb
Selection.CreateAutoTextEntry "handdel", _ 
 "Mailing Instructions"
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]