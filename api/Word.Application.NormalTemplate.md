---
title: Application.NormalTemplate property (Word)
keywords: vbawd10.chm158334984
f1_keywords:
- vbawd10.chm158334984
ms.prod: word
api_name:
- Word.Application.NormalTemplate
ms.assetid: 0ffd1cfd-78da-5189-2504-bbc04bf5b484
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NormalTemplate property (Word)

Returns a  **[Template](Word.Template.md)** object that represents the Normal template.


## Syntax

_expression_. `NormalTemplate`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example inserts the AutoText entry named "Test" from the Normal template, if this entry is contained in the  **AutoTextEntries** collection.


```vb
For Each entry In NormalTemplate.AutoTextEntries 
 If entry.Name = "Test" Then entry.Insert Where:=Selection.Range 
Next entry
```

This example saves the Normal template if changes have been made to it.




```vb
If NormalTemplate.Saved = False Then NormalTemplate.Save
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]