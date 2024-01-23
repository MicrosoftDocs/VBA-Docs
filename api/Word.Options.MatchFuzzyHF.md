---
title: Options.MatchFuzzyHF property (Word)
keywords: vbawd10.chm162988353
f1_keywords:
- vbawd10.chm162988353
api_name:
- Word.Options.MatchFuzzyHF
ms.assetid: fc818d98-8cdc-2dfe-9898-d019a01b2077
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyHF property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Screenshot of the first symbol in the example.](../images/fe283_ZA06051762.gif)
![Screenshot of the second symbol in the example.](../images/fe284_ZA06051763.gif)" and "
![Screenshot of the third symbol in the example.](../images/fe238_ZA06051718.gif)
![Screenshot of the fourth symbol in the example.](../images/fe284_ZA06051763.gif)" and between "
![Screenshot of the fifth symbol in the example.](../images/fe285_ZA06051764.gif)
![Screenshot of the sixth symbol in the example.](../images/fe284_ZA06051763.gif)" and "
![Screenshot of the seventh symbol in the example.](../images/fe267_ZA06051746.gif)
![Screenshot of the eighth symbol in the example.](../images/fe284_ZA06051763.gif)" during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyHF`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Screenshot of symbol #1 in the example.](../images/fe283_ZA06051762.gif)
![Screenshot of symbol #2 in the example.](../images/fe284_ZA06051763.gif)" and "
![Screenshot of symbol #3 in the example.](../images/fe238_ZA06051718.gif)
![Screenshot of symbol #4 in the example.](../images/fe284_ZA06051763.gif)" and between "
![Screenshot of symbol #5 in the example.](../images/fe285_ZA06051764.gif)
![Screenshot of symbol #6 in the example.](../images/fe284_ZA06051763.gif)" and "
![Screenshot of symbol #7 in the example.](../images/fe267_ZA06051746.gif)
![Screenshot of symbol #8 in the example.](../images/fe284_ZA06051763.gif)" during a search.


```vb
Options.MatchFuzzyHF = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]