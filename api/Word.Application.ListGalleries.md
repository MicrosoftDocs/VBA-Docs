---
title: Application.ListGalleries property (Word)
keywords: vbawd10.chm158335041
f1_keywords:
- vbawd10.chm158335041
ms.prod: word
api_name:
- Word.Application.ListGalleries
ms.assetid: 769d3494-3fc3-5a4b-e6d1-a3910107c8bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ListGalleries property (Word)

Returns a  **[ListGalleries](Word.listgalleries.md)** collection that represents the three list template galleries. .


## Syntax

_expression_. `ListGalleries`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Remarks

Each template gallery (Bulleted, Numbered, and Outline Numbered) corresponds to a tab in the  **Bullets and Numbering** dialog box (**Format** menu). For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the variable mylsttmp to the second list template on the  **Outline Numbered** tab in the **Bullets and Numbering** dialog box. The example then applies that template to the first list in the active document.


```vb
Set mylsttmp = _ 
 ListGalleries(wdOutlineNumberGallery).ListTemplates(2) 
ActiveDocument.Lists(1).ApplyListTemplate ListTemplate:=mylsttmp
```

This example cycles through the  **ListGalleries** collection and changes the templates in each list template gallery back to the built-in template.




```vb
For Each listgal In ListGalleries 
 For i = 1 To 7 
 listgal.Reset(i) 
 Next i 
Next listgal
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]