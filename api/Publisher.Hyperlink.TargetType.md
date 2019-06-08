---
title: Hyperlink.TargetType property (Publisher)
keywords: vbapb10.chm4587529
f1_keywords:
- vbapb10.chm4587529
ms.prod: publisher
api_name:
- Publisher.Hyperlink.TargetType
ms.assetid: 1cbc8c36-563c-4464-4f0d-2836682ce532
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlink.TargetType property (Publisher)

Returns a **[PbHlinkTargetType](publisher.pbhlinktargettype.md)** constant that represents the type of hyperlink. Read-only.


## Syntax

_expression_.**TargetType**

_expression_ A variable that represents a **[Hyperlink](Publisher.Hyperlink.md)** object.


## Return value

PbHlinkTargetType


## Remarks

The **TargetType** property value can be one of the **PbHlinkTargetType** constants.

## Example

This example verifies that the specified hyperlink is a URL, and if it is, sets the hyperlink display text and address. This example assumes that there is at least one shape on the first page of the active publication.

```vb
Sub SetHyperlinkTextToDisplay() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Item(1) 
 If .TargetType = pbHlinkTargetTypeURL Then 
 .TextToDisplay = "Tailspin Toys website" 
 .Address = "https://www.tailspintoys.com/" 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]