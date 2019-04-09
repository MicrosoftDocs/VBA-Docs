---
title: Revision.Author property (Word)
keywords: vbawd10.chm159449089
f1_keywords:
- vbawd10.chm159449089
ms.prod: word
api_name:
- Word.Revision.Author
ms.assetid: c56d13d8-e95e-06b7-be83-2df98dbb979c
ms.date: 06/08/2017
localization_priority: Normal
---


# Revision.Author property (Word)

Returns the name of the user who made the specified tracked change. Read-only  **String**.


## Syntax

_expression_. `Author`

_expression_ Required. A variable that represents a '[Revision](Word.Revision.md)' object.


## Example

This example displays the author name for the first tracked change in the first selected section.


```vb
Dim rngSection as Range 
 
Set rngSection = Selection.Sections(1).Range 
MsgBox "Revisions made by " & rngSection.Revisions(1).Author
```


## See also


[Revision Object](Word.Revision.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]