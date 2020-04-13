---
title: Options.PrintDraft property (Word)
keywords: vbawd10.chm162988319
f1_keywords:
- vbawd10.chm162988319
ms.prod: word
api_name:
- Word.Options.PrintDraft
ms.assetid: 23be1e0a-784b-5b0f-107c-78e200e31159
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PrintDraft property (Word)

 **True** if Microsoft Word prints using minimal formatting. Read/write **Boolean**.


## Syntax

_expression_. `PrintDraft`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

Not all printers support draft printing.


## Example

This example sets Word to use draft printing and then prints the active document.


```vb
Options.PrintDraft = True 
ActiveDocument.PrintOut
```

This example returns the current status of the **Draft output** option on the **Print** tab in the **Options** dialog box (**Tools** menu).




```vb
temp = Options.PrintDraft
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]