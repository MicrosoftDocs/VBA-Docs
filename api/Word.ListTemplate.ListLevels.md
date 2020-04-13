---
title: ListTemplate.ListLevels property (Word)
keywords: vbawd10.chm160366594
f1_keywords:
- vbawd10.chm160366594
ms.prod: word
api_name:
- Word.ListTemplate.ListLevels
ms.assetid: ed3c036d-b9be-eeb1-5894-ddf1e2a5f8df
ms.date: 06/08/2017
localization_priority: Normal
---


# ListTemplate.ListLevels property (Word)

Returns a  **[ListLevels](Word.listlevels.md)** collection that represents all the levels for the specified **ListTemplate**.


## Syntax

_expression_. `ListLevels`

 _expression_ An expression that returns a '[ListTemplate](Word.ListTemplate.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the variable myListTemp to the first list template (excluding None) on the **Outline Numbered** tab in the **Bullets and Numbering** dialog box (**Format** menu). Each level in the list has a matching heading style linked to it.


```vb
Set myListTemp = _ 
 ListGalleries(wdOutlineNumberGallery).ListTemplates(1) 
For Each mylevel In myListTemp.ListLevels 
 mylevel.LinkedStyle = "Heading " & mylevel.index 
Next mylevel
```


## See also


[ListTemplate Object](Word.ListTemplate.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]