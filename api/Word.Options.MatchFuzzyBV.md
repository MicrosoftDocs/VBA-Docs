---
title: Options.MatchFuzzyBV property (Word)
keywords: vbawd10.chm162988351
f1_keywords:
- vbawd10.chm162988351
ms.prod: word
api_name:
- Word.Options.MatchFuzzyBV
ms.assetid: 34b82945-06cd-715b-85e3-e09b9f924d84
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.MatchFuzzyBV property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Symbol](../images/fe143_ZA06051648.gif)" and "
![Symbol](../images/fe267_ZA06051746.gif)
![Symbol](../images/fe268_ZA06051747.gif)" and between "
![Symbol](../images/fe278_ZA06051757.gif)" and "
![Symbol](../images/fe238_ZA06051718.gif)
![Symbol](../images/fe268_ZA06051747.gif)" during a search. Read/write  **Boolean**.


## Syntax

_expression_. `MatchFuzzyBV`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Symbol](../images/fe143_ZA06051648.gif)" and "
![Symbol](../images/fe267_ZA06051746.gif)
![Symbol](../images/fe268_ZA06051747.gif)" and between "
![Symbol](../images/fe278_ZA06051757.gif)" and "
![Symbol](../images/fe238_ZA06051718.gif)
![Symbol](../images/fe268_ZA06051747.gif)" during a search.


```vb
Options.MatchFuzzyBV = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]