---
title: Global.ListGalleries property (Word)
keywords: vbawd10.chm163119169
f1_keywords:
- vbawd10.chm163119169
ms.prod: word
api_name:
- Word.Global.ListGalleries
ms.assetid: 56ac5cc2-552a-cff6-95cb-40eebd904eb7
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.ListGalleries property (Word)

Returns a  **ListGalleries** collection that represents the three list template galleries (**Bulleted**,  **Numbered**, and  **Outline Numbered**).


## Syntax

_expression_. `ListGalleries`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


## Remarks

Each gallery corresponds to a tab in the **Bullets and Numbering** dialog box. For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the variable  _mylsttmp_ to the second list template on the **Outline Numbered** tab in the **Bullets and Numbering** dialog box. The example then applies that template to the first list in the active document.


```vb
Set mylsttmp = _ 
 ListGalleries(wdOutlineNumberGallery).ListTemplates(2) 
ActiveDocument.Lists(1).ApplyListTemplate ListTemplate:=mylsttmp
```

This example cycles through the **ListGalleries** collection and changes the templates in each list template gallery back to the built-in template.




```vb
For Each listgal In ListGalleries 
 For i = 1 To 7 
 listgal.Reset(i) 
 Next i 
Next listgal
```


## See also


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]