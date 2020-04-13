---
title: Global.Templates property (Word)
keywords: vbawd10.chm163119171
f1_keywords:
- vbawd10.chm163119171
ms.prod: word
api_name:
- Word.Global.Templates
ms.assetid: 4aa67807-023a-2b52-4773-114d86e340e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Templates property (Word)

Returns a  **Templates** collection that represents all the available templates—global templates and those attached to open documents.


## Syntax

_expression_. `Templates`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the name of each template in the **Templates** collection.


```vb
Count = 1 
For Each aTemplate In Templates 
 MsgBox aTemplate.Name & " is template number " & Count 
 Count = Count + 1 
Next aTemplate
```

In this example, if template one is a global template, its path is stored in  _thePath_. The **ChDir** statement is used to make the folder with the path stored in _thePath_ the current folder. When this change is made, the **Open** dialog box is displayed.




```vb
If Templates(1).Type = wdGlobalTemplate Then 
 thePath = Templates(1).Path 
 If thePath <> "" Then ChDir thePath 
 Dialogs(wdDialogFileOpen).Show 
End If
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]