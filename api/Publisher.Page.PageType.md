---
title: Page.PageType property (Publisher)
keywords: vbapb10.chm393221
f1_keywords:
- vbapb10.chm393221
ms.prod: publisher
api_name:
- Publisher.Page.PageType
ms.assetid: 0bb34de5-ac3e-386c-3b9f-814a476c9695
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.PageType property (Publisher)

Returns a **[PbPageType](publisher.pbpagetype.md)** constant that represents the page type. Read-only.


## Syntax

_expression_.**PageType**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

PbPageType


## Remarks

The **PageType** property value can be one of the **PbPageType** constants declared in the Microsoft Publisher type library.

## Example

This example adds a shape on alternating corners of each page in the active publication.

```vb
Sub GetPageType() 
 Dim pgCount As Page 
 For Each pgCount In ActiveDocument.Pages 
 If pgCount.PageType = pbPageLeftPage Then 
 pgCount.Shapes.AddShape Type:=msoShapeOval, _ 
 Left:=50, Top:=50, Width:=50, Height:=50 
 ElseIf pgCount.PageType = pbPageRightPage Then 
 pgCount.Shapes.AddShape Type:=msoShapeOval, _ 
 Left:=512, Top:=50, Width:=50, Height:=50 
 End If 
 Next 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]