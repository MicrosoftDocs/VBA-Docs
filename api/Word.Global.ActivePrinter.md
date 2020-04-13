---
title: Global.ActivePrinter property (Word)
keywords: vbawd10.chm163119170
f1_keywords:
- vbawd10.chm163119170
ms.prod: word
api_name:
- Word.Global.ActivePrinter
ms.assetid: cf4dcba0-7b26-0569-8ab8-eb641696d0e1
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.ActivePrinter property (Word)

Returns or sets the name of the active printer. Read/write  **String**.


## Syntax

_expression_.**ActivePrinter**

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

Setting the printer using the **ActivePrinter** property changes the default printer. For more information, see [Setting ActivePrinter Changes System Default Printer](https://go.microsoft.com/fwlink/?LinkId=61996).


## Example

This example displays the name of the active printer.


```vb
MsgBox "The name of the active printer is " & ActivePrinter
```

This example makes a network HP LaserJet IIISi printer the active printer.




```vb
Application.ActivePrinter = "HP LaserJet IIISi on \\printers\laser"
```

This example makes a local HP LaserJet 4 printer on LPT1 the active printer.




```vb
Application.ActivePrinter = "HP LaserJet 4 local on LPT1:"
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]