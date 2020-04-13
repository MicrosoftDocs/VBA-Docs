---
title: Sections.Add method (Word)
keywords: vbawd10.chm156893189
f1_keywords:
- vbawd10.chm156893189
ms.prod: word
api_name:
- Word.Sections.Add
ms.assetid: 85063c54-fcd6-8421-2de1-e7fc90289336
ms.date: 06/08/2017
localization_priority: Normal
---


# Sections.Add method (Word)

Returns a  **Section** object that represents a new section added to a document.


## Syntax

_expression_.**Add** (_Range_, _Start_)

_expression_ Required. A variable that represents a '[Sections](Word.sections.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Optional| **Variant**|The range before which you want to insert the section break. If this argument is omitted, the section break is inserted at the end of the document.|
| _Start_|Optional| **Variant**|The type of section break you want to add. Can be one of the **WdSectionStart** constants. If this argument is omitted, a Next Page section break is added.|

## Example

This example adds a Next Page section break before the third paragraph in the active document.


```vb
Set myRange = ActiveDocument.Paragraphs(3).Range 
ActiveDocument.Sections.Add Range:=myRange
```

This example adds a Continuous section break at the selection.




```vb
Set myRange = Selection.Range 
ActiveDocument.Sections.Add Range:=myRange, _ 
 Start:=wdSectionContinuous
```

This example adds a Next Page section break at the end of the active document.




```vb
ActiveDocument.Sections.Add
```


## See also


[Sections Collection Object](Word.sections.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]