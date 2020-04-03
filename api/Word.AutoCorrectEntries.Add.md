---
title: AutoCorrectEntries.Add method (Word)
keywords: vbawd10.chm155713637
f1_keywords:
- vbawd10.chm155713637
ms.prod: word
api_name:
- Word.AutoCorrectEntries.Add
ms.assetid: 670539d8-02f4-dcc9-79bd-20290766b029
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrectEntries.Add method (Word)

Returns an  **AutoCorrectEntry** object that represents a plain-text AutoCorrect entry added to the list of available AutoCorrect entries.


## Syntax

_expression_.**Add** (_Name_, _Value_)

_expression_ Required. A variable that represents an '[AutoCorrectEntries](Word.autocorrectentries.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The text you want to have automatically replaced with the text specified by Value.|
| _Value_|Required| **String**|The text you want to have automatically inserted whenever the text specified by Name is typed.|

## Remarks

Use the  **AddRichText** method to create a formatted AutoCorrect entry.


## Example

This example adds a plain-text AutoCorrect entry for a common misspelling of the word their.


```vb
AutoCorrect.Entries.Add Name:="thier", Value:="their"
```


## See also


[AutoCorrectEntries Collection Object](Word.autocorrectentries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]