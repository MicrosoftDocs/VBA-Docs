---
title: Options.MatchFuzzyZJ property (Word)
keywords: vbawd10.chm162988354
f1_keywords:
- vbawd10.chm162988354
ms.prod: word
api_name:
- Word.Options.MatchFuzzyZJ
ms.assetid: 8f722df0-9fa4-3207-9cad-694cac2d955a
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyZJ property (Word)

 **True** if Microsoft Word ignores the distinction between "
![A screenshot that shows symbol #1 in the example.](../images/fe286_ZA06051765.gif)" and "
![A screenshot that shows symbol #2 in the example.](../images/fe287_ZA06051766.gif)
![A screenshot that shows symbol #3 in the example.](../images/fe209_ZA06051695.gif)" and between "
![A screenshot that shows symbol #4 in the example.](../images/fe288_ZA06051767.gif)" and "
![A screenshot that shows symbol #5 in the example.](../images/fe275_ZA06051754.gif)
![A screenshot that shows symbol #6 in the example.](../images/fe209_ZA06051695.gif)" during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyZJ`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![A screenshot that shows the first symbol in the example.](../images/fe286_ZA06051765.gif)" and "
![A screenshot that shows the second symbol in the example.](../images/fe287_ZA06051766.gif)
![A screenshot that shows the third symbol in the example.](../images/fe209_ZA06051695.gif)" and between "
![A screenshot that shows the fourth symbol in the example.](../images/fe288_ZA06051767.gif)" and "
![A screenshot that shows the fifth symbol in the example.](../images/fe275_ZA06051754.gif)
![A screenshot that shows the sixth symbol in the example.](../images/fe209_ZA06051695.gif)" during a search.


```vb
Options.MatchFuzzyZJ = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]