---
title: Hyperlink.SetPageRelative method (Publisher)
keywords: vbapb10.chm4587542
f1_keywords:
- vbapb10.chm4587542
ms.prod: publisher
api_name:
- Publisher.Hyperlink.SetPageRelative
ms.assetid: 4b2f2e84-09ce-cef6-6f22-b82642cc71fe
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlink.SetPageRelative method (Publisher)

Sets the target type for the specified hyperlink.


## Syntax

_expression_.**SetPageRelative** (_RelativePage_)

_expression_ A variable that represents a **[Hyperlink](Publisher.Hyperlink.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_RelativePage_|Required| **[PbHlinkTargetType](publisher.pbhlinktargettype.md)**|The target type of the hyperlink. Can be one of the **PbHlinkTargetType** constants declared in the Microsoft Publisher type library.|


## Example

The following example adds four new hyperlinks to shape one on page one of the active publication and sets their targets accordingly.

```vb
Sub SetHyperlinkRelativeTarget() 
 Dim hypNew As Hyperlink 
 Dim txtRng As TextRange 
 
 ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox Orientation:=pbTextOrientationHorizontal, _ 
 Left:=10, Top:=10, Width:=200, Height:=200 
 
 Set txtRng = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange 
 
 txtRng.Text = "First Page" & vbCrLf 
 
 Set txtRng = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="https://www.tailspintoys.com/") 
 
 'Change hyperlink to be a Page-relative link 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypeFirstPage 
 
 txtRng.Collapse pbCollapseEnd 
 txtRng.Text = "Previous Page" & vbCrLf 
 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="https://www.tailspintoys.com/") 
 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypePreviousPage 
 
 txtRng.Collapse pbCollapseEnd 
 txtRng.Text = "Next Page" & vbCrLf 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="https://www.tailspintoys.com/") 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypeNextPage 
 
 txtRng.Collapse pbCollapseEnd 
 txtRng.Text = "Last Page" & vbCrLf 
 Set hypNew = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Hyperlinks.Add(Text:=txtRng, _ 
 Address:="https://www.tailspintoys.com/") 
 hypNew.SetPageRelative RelativePage:=pbHlinkTargetTypeLastPage 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]