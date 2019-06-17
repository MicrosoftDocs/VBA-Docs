---
title: WebPageOptions.IncludePageOnNewWebNavigationBars property (Publisher)
keywords: vbapb10.chm544773
f1_keywords:
- vbapb10.chm544773
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.IncludePageOnNewWebNavigationBars
ms.assetid: 5e2f60d0-e812-8ca1-e54b-33a1f9eedf84
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.IncludePageOnNewWebNavigationBars property (Publisher)

Returns or sets a **Boolean** value that specifies whether a link to a webpage will be added to the automatic navigation bars of new pages. Read/write.


## Syntax

_expression_.**IncludePageOnNewWebNavigationBars**

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Return value

Boolean


## Remarks

The default value of the **IncludePageOnNewWebNavigationBars** property is **False**, which means that links to the specified page are not added to the automatic navigation bars of new pages.

Setting this property to **False** does not remove links to the specified page from any automatic navigation bars that already include them, but it does prevent links to the page from being added to automatic navigation bars of new pages.

Setting this property to **True** applies only to automatic navigation bars of new pages, and does not update existing automatic navigation bars within the web publication.

When adding a new page to the web publication by using the **[Pages.Add](Publisher.Pages.Add.md)** method, the optional _AddHyperlinkToWebNavBar_ parameter can be used to specify whether links to the new page are added to existing automatic navigation bars. The value of this parameter is used to populate the value of the **IncludePageOnNewWebNavigationBars** property.

## Example

The following example specifies that links to page two of the active web publication should be added to the automatic navigation bars of new pages. Note that if a new page is added to the publication after this point, the **IncludePageOnNewWebNavigationBars** property will be **False**.

```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
With theWPO 
 .IncludePageOnNewWebNavigationBars = True 
End With
```

<br/>

The following example demonstrates adding two new pages to the publication by using the **Pages.Add** method. The _AddHyperlinkToWebNavBar_ parameter is set to **True**, which specifies that links to these two new pages be added to the automatic navigation bars of existing pages.

Another page is then added to the publication, and the _AddHyperlinkToWebNavBar_ parameter is omitted. This means that the **IncludePageOnNewWebNavigationBars** property is **False** for the newly added page, and links to this page will not be included in the automatic navigation bars of existing pages.

```vb
Dim thePage As page 
Dim thePage2 As page 
 
Set thePage = ActiveDocument.Pages.Add(Count:=2, _ 
 After:=4, AddHyperlinkToWebNavBar:=True) 
 
Set thePage2 = ActiveDocument.Pages.Add(Count:=1, After:=6)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]