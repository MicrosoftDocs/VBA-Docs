---
title: Options.MatchFuzzyDZ property (Word)
keywords: vbawd10.chm162988350
f1_keywords:
- vbawd10.chm162988350
api_name:
- Word.Options.MatchFuzzyDZ
ms.assetid: 4594528b-3855-512d-9738-878ce68c4bf7
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyDZ property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Screenshot of the first symbol in the example.](../images/fe274_ZA06051753.gif)" and "
![Screenshot of the second symbol in the example.](../images/fe275_ZA06051754.gif)" and between "
![Screenshot of the third symbol in the example.](../images/fe276_ZA06051755.gif)" and "
![Screenshot of the fourth symbol in the example.](../images/fe277_ZA06051756.gif)" during a search. Read/write **Boolean**.


## Syntax

_expression_.**MatchFuzzyDZ**

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Screenshot of symbol #1 in the example.](../images/fe274_ZA06051753.gif)" and "
![Screenshot of symbol #2 in the example.](../images/fe275_ZA06051754.gif)" and between "
![Screenshot of symbol #3 in the example.](../images/fe276_ZA06051755.gif)" and "
![Screenshot of symbol #4 in the example.](../images/fe277_ZA06051756.gif)" during a search.


```vb
Options.MatchFuzzyDZ = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]