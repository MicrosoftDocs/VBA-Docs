---
title: Options.SaveInterval property (Word)
keywords: vbawd10.chm162988077
f1_keywords:
- vbawd10.chm162988077
ms.prod: word
api_name:
- Word.Options.SaveInterval
ms.assetid: 0f0ce021-f883-60d3-6dfe-f17c626dd07e
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.SaveInterval property (Word)

Returns or sets the time interval in minutes for saving AutoRecover information. Read/write  **Long**.


## Syntax

_expression_. `SaveInterval`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

Set the **SaveInterval** property to 0 (zero) to turn off saving AutoRecover information.


## Example

This example sets Word to save AutoRecover information for all open documents every five minutes.


```vb
Options.SaveInterval = 5
```

This example prevents Word from saving AutoRecover information.




```vb
Options.SaveInterval = 0
```

This example returns the current status of the **Save AutoRecover info every** option on the **Save** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.SaveInterval
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]