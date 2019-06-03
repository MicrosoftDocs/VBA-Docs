---
title: Global.Dialogs property (Word)
keywords: vbawd10.chm163119123
f1_keywords:
- vbawd10.chm163119123
ms.prod: word
api_name:
- Word.Global.Dialogs
ms.assetid: 7eea3680-b232-c18a-d99a-d7c2a5b29cd4
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Dialogs property (Word)

Returns a  **[Dialogs](Word.dialogs.md)** collection that represents all the built-in dialog boxes in Word. Read-only.


## Syntax

_expression_. `Dialogs`

_expression_ A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the built-in Find dialog box, with "Hello" in the Find What box.


```vb
Dim dlgFind As Dialog 
 
Set dlgFind = Dialogs(wdDialogEditFind) 
 
With dlgFind 
 .Find = "Hello" 
 .Show 
End With
```

This example displays the built-in Open dialog box showing all file types.




```vb
With Dialogs(wdDialogFileOpen) 
 .Name = "*.*" 
 .Show 
End With
```

This example prints the active document, using the settings from the Print dialog box.




```vb
Dialogs(wdDialogFilePrint).Execute
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]