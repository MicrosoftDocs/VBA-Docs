---
title: Sections.PageSetup property (Word)
keywords: vbawd10.chm156894285
f1_keywords:
- vbawd10.chm156894285
ms.prod: word
api_name:
- Word.Sections.PageSetup
ms.assetid: d6d86ddf-bb28-f2fc-49ff-7cfe04853fba
ms.date: 06/08/2017
localization_priority: Normal
---


# Sections.PageSetup property (Word)

Returns a  **PageSetup** object that's associated with the specified document, range, section, sections, or selection.


## Syntax

_expression_.**PageSetup**

_expression_ A variable that represents a '[Sections](Word.sections.md)' collection.


## Example

This example sets the gutter for the first section in Summary.doc to 36 points (0.5 inch).


```vb
Documents("Summary.doc").Sections(1).PageSetup.Gutter = 36
```


## See also


[Sections Collection Object](Word.sections.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]