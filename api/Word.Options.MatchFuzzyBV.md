---
title: Options.MatchFuzzyBV property (Word)
keywords: vbawd10.chm162988351
f1_keywords:
- vbawd10.chm162988351
api_name:
- Word.Options.MatchFuzzyBV
ms.assetid: 34b82945-06cd-715b-85e3-e09b9f924d84
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyBV property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Screenshot of symbol #1 in the example.](../images/fe143_ZA06051648.gif)" and "
![Screenshot of symbol #2 in the example.](../images/fe267_ZA06051746.gif)
![Screenshot of symbol #3 in the example.](../images/fe268_ZA06051747.gif)" and between "
![Screenshot of symbol #4 in the example.](../images/fe278_ZA06051757.gif)" and "
![Screenshot of symbol #5 in the example.](../images/fe238_ZA06051718.gif)
![Screenshot of symbol #6 in the example.](../images/fe268_ZA06051747.gif)" during a search. Read/write **Boolean**.


## Syntax

_expression_.**MatchFuzzyBV**

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Screenshot of the first symbol in the example.](../images/fe143_ZA06051648.gif)" and "
![Screenshot of the second symbol in the example.](../images/fe267_ZA06051746.gif)
![Screenshot of the third symbol in the example.](../images/fe268_ZA06051747.gif)" and between "
![Screenshot of the fourth symbol in the example.](../images/fe278_ZA06051757.gif)" and "
![Screenshot of the fifth symbol in the example.](../images/fe238_ZA06051718.gif)
![Screenshot of the sixth symbol in the example.](../images/fe268_ZA06051747.gif)" during a search.


```vb
Options.MatchFuzzyBV = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]