---
title: HeaderFooter object (Word)
ms.prod: word
api_name:
- Word.HeaderFooter
ms.assetid: 3f2f926a-9220-5536-80ed-af63d2feb016
ms.date: 06/08/2017
localization_priority: Normal
---


# HeaderFooter object (Word)

Represents a single header or footer. The  **HeaderFooter** object is a member of the **[HeadersFooters](Word.headersfooters.md)** collection. The **HeadersFooters** collection includes all headers and footers in the specified document section.


## Remarks

Use  **Headers** (Index) or **Footers** (Index), where index is one of the **WdHeaderFooterIndex** constants (**wdHeaderFooterEvenPages**, **wdHeaderFooterFirstPage**, or **wdHeaderFooterPrimary**), to return a single **HeaderFooter** object. The following example changes the text of both the primary header and the primary footer in the first section of the active document.


```vb
With ActiveDocument.Sections(1) 
 .Headers(wdHeaderFooterPrimary).Range.Text = "Header text" 
 .Footers(wdHeaderFooterPrimary).Range.Text = "Footer text" 
End With
```

You can also return a single  **HeaderFooter** object by using the **HeaderFooter** property with a **Selection** object.


> [!NOTE] 
> You cannot add  **HeaderFooter** objects to the **[HeadersFooters](Word.headersfooters.md)** collection.

Use the  **DifferentFirstPageHeaderFooter** property with the **PageSetup** object to specify a different first page. The following example inserts text into the first page footer in the active document.




```vb
With ActiveDocument 
 .PageSetup.DifferentFirstPageHeaderFooter = True 
 .Sections(1).Footers(wdHeaderFooterFirstPage) _ 
 .Range.InsertBefore _ 
 "Written by Joe Smith" 
End With
```

Use the  **OddAndEvenPagesHeaderFooter** property with the **PageSetup** object to specify different odd and even page headers and footers. If the **OddAndEvenPagesHeaderFooter** property is **True**, you can return an odd header or footer by using **wdHeaderFooterPrimary**, and you can return an even header or footer by using **wdHeaderFooterEvenPages**.

Use the  **Add** method with the **PageNumbers** object to add a page number to a header or footer. The following example adds page numbers to the primary footer in the first section of the active document.




```vb
With ActiveDocument.Sections(1) 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add 
End With
```


## Properties



|Name|
|:-----|
|[Application](Word.HeaderFooter.Application.md)|
|[Creator](Word.HeaderFooter.Creator.md)|
|[Exists](Word.HeaderFooter.Exists.md)|
|[Index](Word.HeaderFooter.Index.md)|
|[IsHeader](Word.HeaderFooter.IsHeader.md)|
|[LinkToPrevious](Word.HeaderFooter.LinkToPrevious.md)|
|[PageNumbers](Word.HeaderFooter.PageNumbers.md)|
|[Parent](Word.HeaderFooter.Parent.md)|
|[Range](Word.HeaderFooter.Range.md)|
|[Shapes](Word.HeaderFooter.Shapes.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
