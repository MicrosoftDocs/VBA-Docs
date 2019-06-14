---
title: TextRange.InsertPageNumber method (Publisher)
keywords: vbapb10.chm5308486
f1_keywords:
- vbapb10.chm5308486
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertPageNumber
ms.assetid: f71d3b40-0263-93fa-d7e3-d815b90f71f7
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.InsertPageNumber method (Publisher)

Returns a **TextRange** object that represents a page number field in a publication.


## Syntax

_expression_.**InsertPageNumber** (_Type_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_|Optional| **[PbPageNumberType](publisher.pbpagenumbertype.md)** |Specifies whether the page number is the current page number or the next or previous page number of a linked text box. Can be one of the **PbPageNumberType** constants.|

## Return value

TextRange


## Example

This example inserts a page number field in a shape on the master page so that the current page number appears at the top of each page.

```vb
Sub PageNumberShape() 
 With ActiveDocument.MasterPages(1).Shapes _ 
 .AddShape(Type:=msoShape5pointStar, Left:=36, _ 
 Top:=36, Width:=50, Height:=50) 
 With .TextFrame.TextRange 
 .InsertPageNumber 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 End With 
 .Fill.ForeColor.RGB = RGB(Red:=125, Green:=125, Blue:=255) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]