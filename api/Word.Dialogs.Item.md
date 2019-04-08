---
title: Dialogs.Item method (Word)
keywords: vbawd10.chm152043520
f1_keywords:
- vbawd10.chm152043520
ms.prod: word
api_name:
- Word.Dialogs.Item
ms.assetid: 8a7826ce-a5b9-e0af-29cb-5dea299ab266
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialogs.Item method (Word)

Returns a dialog in Microsoft Word.


## Syntax

_expression_.**Item** (_Index_)

_expression_ Required. A variable that represents a '[Dialogs](Word.dialogs.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **WdWordDialog**|A constant that specifies the dialog.|

## Return value

Dialog


## Example

This example displays the Page Setup dialog.


```vb
Sub DialogItem() 
 Application.Dialogs.Item(wdDialogFileDocumentLayout).Display 
End Sub
```


## See also


[Dialogs Collection Object](Word.dialogs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]