---
title: Pages.Add method (Publisher)
keywords: vbapb10.chm458757
f1_keywords:
- vbapb10.chm458757
ms.prod: publisher
api_name:
- Publisher.Pages.Add
ms.assetid: 3c22aa15-c1dc-94c8-62d6-a1bc9635cd89
ms.date: 06/12/2019
localization_priority: Normal
---


# Pages.Add method (Publisher)

Adds a new **[Page](publisher.page.md)** object to the specified **Pages** object and returns the new **Page** object.


## Syntax

_expression_.**Add** (_Count_, _After_, _DuplicateObjectsOnPage_, _AddHyperlinkToWebNavBar_)

_expression_ A variable that represents a **[Pages](Publisher.Pages.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Count_|Required| **Long**|The number of new pages to add.|
|_After_|Required| **Long**|The page index of the page after which to add the new pages. A zero for this argument adds new pages at the beginning of the publication.|
| _DuplicateObjectsOnPage_|Optional| **Long**|The page index of the page from which objects should be copied to the new pages. If this argument is omitted, the new pages will be blank. The default is -1: do not duplicate pages.|
|_AddHyperlinkToWebNavBar_|Optional| **Boolean**|Specifies whether links to the new pages are added to the automatic navigation bars of existing pages. If **True**, links to the new pages are added to the automatic navigation bars of existing pages only.<br/><br/> If **False**, links to the new pages are not added to the automatic navigation bars of existing pages or new pages added in the future. The default is **False**.|

## Return value

Page


## Example

The following example adds four new pages after the first page in the publication, and copies all the objects from the first page to the new pages.

```vb
Dim pgNew As Page 
 
Set pgNew = ActiveDocument.Pages _ 
 .Add(Count:=4, After:=1, DuplicateObjectsOnPage:=1)
```

<br/>

The following example demonstrates adding two new pages to the publication and setting the _AddHyperlinkToWebNavBar_ parameter to **True** for these two pages. This specifies that links to these two new pages be added to the automatic navigation bars of existing pages and those added in the future.

Another page is then added to the publication, and the _AddHyperlinkToWebNavBar_ is omitted. This means that the **[IncludePageOnNewWebNavigationBars](publisher.webpageoptions.includepageonnewwebnavigationbars.md)** property is **False** for the newly added page, and links to this page are not included in the automatic navigation bars of existing pages.

```vb
Dim thePage As page 
Dim thePage2 As page 
 
Set thePage = ActiveDocument.Pages.Add(Count:=2, _ 
 After:=4, AddHyperlinkToWebNavBar:=True) 
 
Set thePage2 = ActiveDocument.Pages.Add(Count:=1, After:=6)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]