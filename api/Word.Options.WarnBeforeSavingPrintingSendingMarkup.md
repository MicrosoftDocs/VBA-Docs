---
title: Options.WarnBeforeSavingPrintingSendingMarkup property (Word)
keywords: vbawd10.chm162988478
f1_keywords:
- vbawd10.chm162988478
ms.prod: word
api_name:
- Word.Options.WarnBeforeSavingPrintingSendingMarkup
ms.assetid: 3d507ad6-5d8f-f20e-eefe-2499f0507b6f
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.WarnBeforeSavingPrintingSendingMarkup property (Word)

 **True** for Microsoft Word to display a warning when saving, printing, or sending as email a document containing comments or tracked changes. Read/write **Boolean**.


## Syntax

_expression_. `WarnBeforeSavingPrintingSendingMarkup`

 _expression_ An expression that returns a '[Options](Word.Options.md)' object.


## Example

This example prints the active document but allows the user to stop the print if the document contains tracked changes or comments.


```vb
Sub SaferPrint 
 Dim blnOldState as Boolean 
 
 'Save old state in variable 
 blnOldState = Application.Options.WarnBeforeSavingPrintingSendingMarkup 
 
 'Turn on warning 
 Application.Options.WarnBeforeSavingPrintingSendingMarkup = True 
 
 'Print document 
 ActiveDocument.PrintOut 
 
 'Restore original warning state 
 Application.Options.WarnBeforeSavingPrintingSendingMarkup = blnOldState 
 
EndSub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]