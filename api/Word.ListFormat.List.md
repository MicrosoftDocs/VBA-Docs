---
title: ListFormat.List property (Word)
keywords: vbawd10.chm163577925
f1_keywords:
- vbawd10.chm163577925
ms.prod: word
api_name:
- Word.ListFormat.List
ms.assetid: e320f0b9-d19c-34d4-b215-395312eadf73
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.List property (Word)

Returns a  **[List](Word.List.md)** object that represents the first formatted list contained in the specified **ListFormat** object.


## Syntax

_expression_.**List**

 _expression_ An expression that returns a '[ListFormat](Word.ListFormat.md)' object.


## Remarks

If the first paragraph in the range for the  **ListFormat** object is not formatted as a list, the **List** property returns nothing.


## Example

This example returns the first list in the selection, and then it applies the first list template (excluding None) on the  **Numbered** tab in the **Bullets and Numbering** dialog box (**Format** menu). The selection can only contain one list.


```vb
Set mylist = Selection.Range.ListFormat.List 
mylist.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdNumberGallery) _ 
 .ListTemplates(1)
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]