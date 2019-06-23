---
title: Application.Dialogs property (Word)
keywords: vbawd10.chm158334995
f1_keywords:
- vbawd10.chm158334995
ms.prod: word
api_name:
- Word.Application.Dialogs
ms.assetid: 17acdfab-32d2-ddb8-04aa-692f9ffb20b8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Dialogs property (Word)

Returns a  **[Dialogs](Word.dialogs.md)** collection that represents all the built-in dialog boxes in Word. Read-only.


## Syntax

_expression_. `Dialogs`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md). 

For a list of built-in dialog boxes, see the **[WdWordDialog](Word.WdWordDialog.md)** enumeration.


## Example

This example displays the built-in  **Find** dialog box, with "Hello" in the **Find What** box.


```vb
Dim dlgFind As Dialog 
 
Set dlgFind = Dialogs(wdDialogEditFind) 
 
With dlgFind 
 .Find = "Hello" 
 .Show 
End With
```

This example displays the built-in  **Open** dialog box showing all file types.




```vb
With Dialogs(wdDialogFileOpen) 
 .Name = "*.*" 
 .Show 
End With
```

This example prints the active document, using the settings from the  **Print** dialog box.




```vb
Dialogs(wdDialogFilePrint).Execute
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]