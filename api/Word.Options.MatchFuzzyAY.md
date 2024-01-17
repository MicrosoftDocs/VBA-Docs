---
title: Options.MatchFuzzyAY property (Word)
keywords: vbawd10.chm162988355
f1_keywords:
- vbawd10.chm162988355
api_name:
- Word.Options.MatchFuzzyAY
ms.assetid: f9a56522-f3a8-0527-e0e9-9144ccc468bc
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyAY property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Screenshot of the first symbol in the example.](../images/fe289_ZA06051768.gif)" and "
![Screenshot of the second symbol in the example.](../images/fe241_ZA06051721.gif)" following 
![Screenshot of the third symbol in the example.](../images/fe144_ZA06051649.gif)-row and 
![Screenshot of the fourth symbol in the example.](../images/fe209_ZA06051695.gif)-row characters during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyAY`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Screenshot of symbol #1 in the example.](../images/fe289_ZA06051768.gif)" and "
![Screenshot of symbol #2 in the example.](../images/fe241_ZA06051721.gif)" following 
![Screenshot of symbol #3 in the example.](../images/fe144_ZA06051649.gif)-row and 
![Screenshot of symbol #4 in the example.](../images/fe209_ZA06051695.gif)-row characters during a search.


```vb
Options.MatchFuzzyAY = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]