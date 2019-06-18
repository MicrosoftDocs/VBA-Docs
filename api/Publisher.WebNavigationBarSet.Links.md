---
title: WebNavigationBarSet.Links property (Publisher)
keywords: vbapb10.chm8519697
f1_keywords:
- vbapb10.chm8519697
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.Links
ms.assetid: 9f155781-390b-ad77-8db7-5099be1409ce
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.Links property (Publisher)

Returns a **[WebNavigationBarHyperlinks](publisher.webnavigationbarhyperlinks.md)** collection containing all of the hyperlinks in the specified web navigation bar set. Read/write.


## Syntax

_expression_.**Links**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

WebNavigationBarHyperlinks


## Example

This example returns the web navigation bar hyperlinks of the first web navigation bar set of the active document.

```vb
ActiveDocument.WebNavigationBarSets(1).Links
```

<br/>

The following example adds a new web navigation bar set to the active document, adds a hyperlink to the navigation bar, and then adds the navigation bar to every page of the publication that has the _AddHyperlinkToWebNavBar_ parameter (**[Pages.Add](publisher.pages.add.md)** method) set to **True** or the **[WebPageOptions.IncludePageOnNewWebNavigationBars](publisher.webpageoptions.includepageonnewwebnavigationbars.md)** property set to **True**.

```vb
With ActiveDocument.WebNavigationBarSets.AddSet(Name:="WebNavigationBarSet1") 
 With .Links 
 .Add Address:="www.microsoft.com", TextToDisplay:="Microsoft", Index:=1 
 End With 
 .AddToEveryPage Left:=10, Top:=10 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]