---
title: ReaderSpread Object (Publisher)
keywords: vbapb10.chm589823
f1_keywords:
- vbapb10.chm589823
ms.prod: publisher
api_name:
- Publisher.ReaderSpread
ms.assetid: 32c55e79-2217-654f-730c-9abaa2cfb9de
ms.date: 06/08/2017
---


# ReaderSpread Object (Publisher)

Represents the reader spread (not the printer spread) for the page. A reader spread generally contains one or two pages. The  **ReaderSpread** object properties provide information about whether pages are facing and how those pages are laid out. For example, in facing page view, pages two and three can be side-by-side or one on top of the other.
 


## Example

Use the  **[ReaderSpread](Publisher.Page.ReaderSpread.md)** property to access the **ReaderSpread** object for a page. Use the **[PageCount](Publisher.ReaderSpread.PageCount.md)** property to determine if the reader spread includes one page or two facing pages. This example checks to see if the reader spread includes less than two pages. If it does, it changes the reader spread to include two pages.
 

 

```
Sub SetFacingPages() 
 With ActiveDocument 
 If .Pages.Count >= 2 Then 
 If .Pages(2).ReaderSpread.PageCount < 2 Then _ 
 .ViewTwoPageSpread = True 
 End If 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](Publisher.ReaderSpread.Application.md)|
|[Height](Publisher.ReaderSpread.Height.md)|
|[Left](Publisher.ReaderSpread.Left.md)|
|[PageCount](Publisher.ReaderSpread.PageCount.md)|
|[Pages](Publisher.ReaderSpread.Pages.md)|
|[Parent](Publisher.ReaderSpread.Parent.md)|
|[Top](Publisher.ReaderSpread.Top.md)|
|[Width](readerspread-width-property-publisher.md)|

