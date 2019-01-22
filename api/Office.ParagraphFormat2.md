---
title: ParagraphFormat2 object (Office)
ms.prod: office
api_name:
- Office.ParagraphFormat2
ms.assetid: 05ff2b24-9603-f923-d053-e736fb2ba389
ms.date: 01/22/2019
localization_priority: Normal
---


# ParagraphFormat2 object (Office)

Represents the paragraph formatting of a text range.


## Example

The following example left aligns the paragraphs in shape two on slide one in the active PowerPoint presentation.


```vb
ActivePresentation.Slides(1).Shapes(2).TextFrame2.TextRange2 _ 
 .ParagraphFormat2.Alignment = ppAlignLeft 

```


## See also

- [ParagraphFormat2 object members](overview/library-reference/paragraphformat2-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]