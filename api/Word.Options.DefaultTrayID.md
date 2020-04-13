---
title: Options.DefaultTrayID property (Word)
keywords: vbawd10.chm162988072
f1_keywords:
- vbawd10.chm162988072
ms.prod: word
api_name:
- Word.Options.DefaultTrayID
ms.assetid: 3a6c265b-f178-318b-bd29-944873c6b036
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultTrayID property (Word)

Returns or sets the default tray your printer uses to print documents. Read/write  **WdPaperTray**.


## Syntax

_expression_. `DefaultTrayID`

_expression_ A variable that represents an **[Options](Word.Options.md)** object.


## Remarks

You can use the **DefaultTray**property with a string from the **Default tray** box on the **Print** tab in the **Options** dialog box to set this same option.


## Example

This example sets Word to use the upper print tray, and then it prints the active document.


```vb
Options.DefaultTrayID = wdPrinterUpperBin 
ActiveDocument.PrintOut
```

This example returns the current setting of the **Default** tray option on the **Print** tab in the **Options** dialog box.




```vb
Dim lngTray As Long 
 
lngTray = Options.DefaultTrayID
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]