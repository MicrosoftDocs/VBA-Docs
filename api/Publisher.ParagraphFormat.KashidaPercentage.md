---
title: ParagraphFormat.KashidaPercentage property (Publisher)
keywords: vbapb10.chm5439513
f1_keywords:
- vbapb10.chm5439513
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.KashidaPercentage
ms.assetid: d62aa512-cce6-2e78-657f-51ff1b2cbcf8
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.KashidaPercentage property (Publisher)

Returns or sets a **Long** indicating the percentage by which kashidas are to be lengthened for the specified paragraphs. Valid values are from 0 to 100. Read/write.


## Syntax

_expression_.**KashidaPercentage**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

Long


## Remarks

The **[Alignment](Publisher.ParagraphFormat.Alignment.md)** property of the specified paragraphs must be set to the **[PbParagraphAlignmentType](publisher.pbparagraphalignmenttype.md)** enumerations's **pbParagraphAlignmentKashida** constant, or the **KashidaPercentage** property is ignored.


## Example

The following example sets the paragraphs in shape one on page one of the active publication to kashida alignment and specifies that kashidas are to be lengthened by 20 percent.

```vb
With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
 .Alignment = pbParagraphAlignmentKashida 
 .KashidaPercentage = 20 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]