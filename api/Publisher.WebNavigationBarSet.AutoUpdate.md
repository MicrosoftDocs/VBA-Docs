---
title: WebNavigationBarSet.AutoUpdate property (Publisher)
keywords: vbapb10.chm8519689
f1_keywords:
- vbapb10.chm8519689
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.AutoUpdate
ms.assetid: b9ce8dde-c09f-6fe9-6935-cb4903a17b85
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.AutoUpdate property (Publisher)

**True** if all pages are added to the specified web navigation bar set and that adding new pages updates the navigation bar with a corresponding item. 

Pages must have the **AddHyperlinkToWebNavbar** set to **True** or the **[WebPageOptions.IncludePageOnNewWebNavigationBars](publisher.webpageoptions.includepageonnewwebnavigationbars.md)** property set to **True** to be added or updated within the specified **WebNavigationBarSet**. Read/write **Boolean**.


## Syntax

_expression_.**AutoUpdate**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

Boolean


## Remarks

This property determines whether the existing pages in the publication are added to the navigation bar and if added pages are also updated. These pages must be marked with the **AddHyperlinkToWebNavbar** set to **True** or the **WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True** to be added or updated within the specified **WebNavigationBarSet**. 

Changing this setting does not change the number of items in the bar; it just determines whether new pages are added. By setting this value to **False**, it is possible to design specific navigation bars for specific content pages in a website that do not contain all the available hyperlinks in the publication.

The default value is **True**. 


## Example

The following example adds a new web navigation bar set to the active document, with text style buttons and auto update set to **False** so that page links are not added, or new pages are not automatically updated in the navigation bar. The web navigation bar is then added to the first page of the publication.

```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newBar") 
With objWebNav 
 .AutoUpdate = False 
 .ButtonStyle = pbnbButtonStyleText 
End With 
ActiveDocument.Pages(1).Shapes.AddWebNavigationBar _ 
 Name:="newBar", Left:=10, Top:=10 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]