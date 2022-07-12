---
title: Options.MatchFuzzyHF property (Word)
keywords: vbawd10.chm162988353
f1_keywords:
- vbawd10.chm162988353
ms.prod: word
api_name:
- Word.Options.MatchFuzzyHF
ms.assetid: fc818d98-8cdc-2dfe-9898-d019a01b2077
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyHF property (Word)

 **True** if Microsoft Word ignores the distinction between "
![A screenshot that shows the first symbol in the example.](../images/fe283_ZA06051762.gif)
![A screenshot that shows the second symbol in the example.](../images/fe284_ZA06051763.gif)" and "
![A screenshot that shows the third symbol in the example.](../images/fe238_ZA06051718.gif)
![A screenshot that shows the fourth symbol in the example.](../images/fe284_ZA06051763.gif)" and between "
![A screenshot that shows the fifth symbol in the example.](../images/fe285_ZA06051764.gif)
![A screenshot that shows the sixth symbol in the example.](../images/fe284_ZA06051763.gif)" and "
![A screenshot that shows the seventh symbol in the example.](../images/fe267_ZA06051746.gif)
![A screenshot that shows the eighth symbol in the example.](../images/fe284_ZA06051763.gif)" during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyHF`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![A screenshot that shows symbol #1 in the example.](../images/fe283_ZA06051762.gif)
![A screenshot that shows symbol #2 in the example.](../images/fe284_ZA06051763.gif)" and "
![A screenshot that shows symbol #3 in the example.](../images/fe238_ZA06051718.gif)
![A screenshot that shows symbol #4 in the example.](../images/fe284_ZA06051763.gif)" and between "
![A screenshot that shows symbol #5 in the example.](../images/fe285_ZA06051764.gif)
![A screenshot that shows symbol #6 in the example.](../images/fe284_ZA06051763.gif)" and "
![A screenshot that shows symbol #7 in the example.](../images/fe267_ZA06051746.gif)
![A screenshot that shows symbol #8 in the example.](../images/fe284_ZA06051763.gif)" during a search.


```vb
Options.MatchFuzzyHF = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]