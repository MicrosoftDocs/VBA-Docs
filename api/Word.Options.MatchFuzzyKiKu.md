---
title: Options.MatchFuzzyKiKu property (Word)
keywords: vbawd10.chm162988356
f1_keywords:
- vbawd10.chm162988356
api_name:
- Word.Options.MatchFuzzyKiKu
ms.assetid: 2e0bde64-f8c2-c61d-1cb3-b8ee3fa8d22d
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyKiKu property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Screenshot of symbol #1 in the example.](../images/fe107_ZA06051631.gif)" and "
![Screenshot of symbol #2 in the example.](../images/fe112_ZA06051635.gif)" before 
![Screenshot of symbol #3 in the example.](../images/fe290_ZA06051769.gif)-row characters during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyKiKu`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Screenshot of the first symbol in the example.](../images/fe107_ZA06051631.gif)" and "
![Screenshot of the second symbol in the example.](../images/fe112_ZA06051635.gif)" before 
![Screenshot of the third symbol in the example.](../images/fe290_ZA06051769.gif)-row characters during a search.


```vb
Options.MatchFuzzyKiKu = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]