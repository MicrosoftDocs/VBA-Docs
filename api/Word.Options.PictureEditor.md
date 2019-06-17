---
title: Options.PictureEditor property (Word)
keywords: vbawd10.chm162988105
f1_keywords:
- vbawd10.chm162988105
ms.prod: word
api_name:
- Word.Options.PictureEditor
ms.assetid: bdf0435b-c0dc-d299-a745-1102e4c46801
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.PictureEditor property (Word)

Returns or sets the name of the application to use to edit pictures. Read/write  **String**.


## Syntax

_expression_. `PictureEditor`

 _expression_ An expression that returns an **[Options](Word.Options.md)** object.


## Remarks

You must use the exact wording displayed in the Picture editor box on the  **Edit** tab of the **Options** dialog box (**Tools** menu). Otherwise, the default setting "Microsoft Word" is used.

If the name of your graphics application doesn't appear in the list, contact the manufacturer of the graphics application for instructions.


## Example

This example sets the application used to edit pictures.


```vb
Options.PictureEditor = "Microsoft Word"
```

This example returns the name of the application to use to edit pictures.




```vb
MsgBox Options.PictureEditor
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]