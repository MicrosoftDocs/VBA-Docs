---
title: TextColumns.SetCount method (Word)
keywords: vbawd10.chm158531786
f1_keywords:
- vbawd10.chm158531786
ms.prod: word
api_name:
- Word.TextColumns.SetCount
ms.assetid: 59ff1b21-5bec-982d-a2b5-7a8d7dc08f9a
ms.date: 06/08/2017
localization_priority: Normal
---


# TextColumns.SetCount method (Word)

Arranges text into the specified number of text columns.


## Syntax

_expression_. `SetCount`( `_NumColumns_` )

_expression_ Required. A variable that represents a '[TextColumns](Word(textcolumns).md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumColumns_|Required| **Long**|The number of columns the text is to be arranged into.|

## Remarks

You can also use the **[Add](Word.TextColumns.Add.md)** method to add a single column to the **TextColumns** collection.


## Example

This example arranges the text in the active document into two columns of equal width.


```vb
ActiveDocument.PageSetup.TextColumns.SetCount NumColumns:=2
```

This example arranges the text in the first section of Brochure.doc into three columns of equal width.




```vb
Documents("Brochure.doc").Sections(1) _ 
 .PageSetup.TextColumns.SetCount NumColumns:=3
```


## See also


[TextColumns Collection Object](Word(textcolumns).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]