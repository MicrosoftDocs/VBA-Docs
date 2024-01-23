---
title: Options.MatchFuzzyTC property (Word)
keywords: vbawd10.chm162988352
f1_keywords:
- vbawd10.chm162988352
api_name:
- Word.Options.MatchFuzzyTC
ms.assetid: 9dc9eb01-d530-f2ac-0bb7-27630ca3ad60
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Options.MatchFuzzyTC property (Word)

 **True** if Microsoft Word ignores the distinction between "
![Screenshot of the first symbol in the example.](../images/fe279_ZA06051758.gif)
![Screenshot of the second symbol in the example.](../images/fe280_ZA06051759.gif)", "
![Screenshot of the third symbol in the example.](../images/fe281_ZA06051760.gif)
![Screenshot of the fourth symbol in the example.](../images/fe280_ZA06051759.gif)", and "
![Screenshot of the fifth symbol in the example.](../images/fe208_ZA06051694.gif)", and between "
![Screenshot of the sixth symbol in the example.](../images/fe282_ZA06051761.gif)
![Screenshot of the seventh symbol in the example.](../images/fe280_ZA06051759.gif)" and "
![Screenshot of the eighth symbol in the example.](../images/fe275_ZA06051754.gif)" during a search. Read/write **Boolean**.


## Syntax

_expression_. `MatchFuzzyTC`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Example

This example sets Microsoft Word to ignore the distinction between "
![Screenshot of symbol #1 in the example.](../images/fe279_ZA06051758.gif)
![Screenshot of symbol #2 in the example.](../images/fe280_ZA06051759.gif)", "
![Screenshot of symbol #3 in the example.](../images/fe281_ZA06051760.gif)
![Screenshot of symbol #4 in the example.](../images/fe280_ZA06051759.gif)", and "
![Screenshot of symbol #5 in the example.](../images/fe208_ZA06051694.gif)", and between "
![Screenshot of symbol #6 in the example.](../images/fe282_ZA06051761.gif)
![Screenshot of symbol #7 in the example.](../images/fe280_ZA06051759.gif)" and "
![Screenshot of symbol #8 in the example.](../images/fe275_ZA06051754.gif)" during a search.


```vb
Options.MatchFuzzyTC = True
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]