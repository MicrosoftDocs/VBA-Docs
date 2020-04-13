---
title: Page object (Word)
keywords: vbawd10.chm169
f1_keywords:
- vbawd10.chm169
ms.prod: word
api_name:
- Word.Page
ms.assetid: 3a3d480a-3876-515f-d13f-7ec23818245f
ms.date: 06/08/2017
localization_priority: Normal
---


# Page object (Word)

Represents a page in a document. Use the **Page** object and the related methods and properties for programmatically defining page layout in a document.


## Remarks

Use the **Item** method to access a specific page in a document. The following example accesses the first page in the active document.


```vb
Dim objPage As Page 
 
Set objPage = ActiveDocument.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```

To access the page number, use the **Information** property of a **Range** or **Selection** object, or the **PageIndex** property of a **Break** object that belongs to the **Breaks** collection of the specified **Page** object.

The **Top** and **Left** properties of the **Page** object always return 0 (zero) indicating the upper-left corner of the page. The **Height** and **Width** properties return the height and width in points (72 points = 1 inch) of the paper size specified in the Page Setup dialog or through the **PageSetup** object. For example, for an 8-1/2 by 11 inch page in portrait mode, the **Height** property returns 792 and the **Width** property returns 612. All four of these properties are read-only.


## Methods



|Name|
|:-----|
|[SaveAsPNG](overview/Word.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Page.Application.md)|
|[Breaks](Word.Page.Breaks.md)|
|[Creator](Word.Page.Creator.md)|
|[EnhMetaFileBits](Word.Page.EnhMetaFileBits.md)|
|[Height](Word.Page.Height.md)|
|[Left](Word.Page.Left.md)|
|[Parent](Word.Page.Parent.md)|
|[Rectangles](Word.Page.Rectangles.md)|
|[Top](Word.Page.Top.md)|
|[Width](Word.Page.Width.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
