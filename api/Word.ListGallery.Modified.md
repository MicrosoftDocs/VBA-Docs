---
title: ListGallery.Modified property (Word)
keywords: vbawd10.chm160694373
f1_keywords:
- vbawd10.chm160694373
ms.prod: word
api_name:
- Word.ListGallery.Modified
ms.assetid: c5acfd0e-5a6f-237e-0a9e-962525fd17d2
ms.date: 06/08/2017
localization_priority: Normal
---


# ListGallery.Modified property (Word)

 **True** if the specified list template is not the built-in list template for that position in the list gallery. Read-only **Boolean**.


## Syntax

_expression_. `Modified` (_Index_)

 _expression_ An expression that returns a '[ListGallery](Word.ListGallery.md)' object.


## Remarks

Use the  **[Reset](Word.ListGallery.Reset.md)** method to set a list template in a list gallery back to the built-in list template.


## Example

This example checks to see whether the first template on the  **Bulleted** tab in the **Bullets and Numbering** dialog box has been changed. If it has, the list template is reset.


```vb
temp = ListGalleries(wdBulletGallery).Modified(1) 
If temp = True Then 
 ListGalleries(wdBulletGallery).Reset(1) 
Else 
 Msgbox "This is the built-in list template." 
End If
```


## See also


[ListGallery Object](Word.ListGallery.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]