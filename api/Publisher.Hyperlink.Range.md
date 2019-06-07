---
title: Hyperlink.Range property (Publisher)
keywords: vbapb10.chm4587526
f1_keywords:
- vbapb10.chm4587526
ms.prod: publisher
api_name:
- Publisher.Hyperlink.Range
ms.assetid: ff105ffe-cb48-0f6a-99ff-eaac0500938f
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlink.Range property (Publisher)

Returns a **[TextRange](Publisher.TextRange.md)** object representing the base text to which the specified hyperlink has been applied.


## Syntax

_expression_.**Range**

_expression_ A variable that represents a **[Hyperlink](Publisher.Hyperlink.md)** object.


## Remarks

If the **Type** property of the specified **Hyperlink** object is a value other than **msoHyperlinkRange**, the **Range** property returns nothing.


## Example

The following example returns the text range associated with the first hyperlink on page one of the active publication and changes the base text to "Go here."

```vb
Dim txtHyperlink As TextRange 
 
txtHyperlink = ActiveDocument.Pages(1) _ 
 .Shapes(1).Hyperlink.Range 
 
txtHyperlink.Text = "Go here"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]