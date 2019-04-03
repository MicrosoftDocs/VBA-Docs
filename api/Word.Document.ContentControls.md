---
title: Document.ContentControls property (Word)
keywords: vbawd10.chm158007804
f1_keywords:
- vbawd10.chm158007804
ms.prod: word
api_name:
- Word.Document.ContentControls
ms.assetid: 86b5af56-3ab4-2440-237e-42af398b260a
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ContentControls property (Word)

Returns a  **[ContentControls](Word.ContentControls.md)** collection that represents all the content controls in a document. Read-only.


## Syntax

_expression_. `ContentControls`

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Example

The following example inserts a drop-down list content control into the active document.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]