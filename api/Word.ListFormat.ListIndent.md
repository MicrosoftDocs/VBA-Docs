---
title: ListFormat.ListIndent method (Word)
keywords: vbawd10.chm163578067
f1_keywords:
- vbawd10.chm163578067
ms.prod: word
api_name:
- Word.ListFormat.ListIndent
ms.assetid: 2c75e457-75f7-378c-b41f-23eb7f6b73da
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.ListIndent method (Word)

Increases the list level of the paragraphs in the range for the specified  **ListFormat** object, in increments of one level.


## Syntax

_expression_. `ListIndent`

_expression_ Required. A variable that represents a '[ListFormat](Word.ListFormat.md)' object.


## Example

This example indents each paragraph in the first list in document one by one level.


```vb
Documents(1).Lists(1).Range.ListFormat.ListIndent
```

This example formats paragraphs four through eight in the active document as an outline-numbered list, and then it indents the paragraphs one level.




```vb
Dim docActive As Document 
Dim rngTemp As Range 
 
Set docActive = ActiveDocument
```




```vb
Set rngTemp = _ 
 docActive.Range( _ 
 Start:=docActive.Paragraphs(4).Range.Start, _ 
 End:=docActive.Paragraphs(8).Range.End) 
 
With rngTemp.ListFormat 
 .ApplyOutlineNumberDefault 
 .ListIndent 
End With
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]