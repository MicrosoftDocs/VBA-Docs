---
title: WebNavigationBarSet Object (Publisher)
keywords: vbapb10.chm8585215
f1_keywords:
- vbapb10.chm8585215
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet
ms.assetid: 03b31cc1-5b24-1a16-710c-73755298066e
ms.date: 06/08/2017
localization_priority: Normal
---


# WebNavigationBarSet Object (Publisher)

Represents a Web navigation bar set for the current document. The  **WebNavigationBarSet** object is a member of the **WebNavigationBarSets** collection, which includes all of the Web navigation bar sets in the current document.
 


## Example

Use  **WebNavigationBarSet**. **AddToEveryPage** (Left, Top, [Width]), where Left is the position of the left edge of the shape, Top is the position of the top edge of the shape, and Width is the width of the shape representing the Web navigation bar set, to add the specified Web navigation bar to every page of a document. The following example adds the first Web navigation bar set to every page that has the **AddHyperlinkToWebNavbar** property set to **True** when adding the page or the **Page.WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True**.
 

 

```vb
Dim objWebNavBarSet as WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets(1) 
objWebNavBarSet.AddToEveryPage Left:=50, Top:=10, Width:=500
```

Use  **WebNavigationBarSet**. **DeleteSetAndInstances** to remove the Web navigation bar set and every instance of it from the document. The following example deletes all instances of each **WebNavigationBarSet** object in the **WebNavigationBarSets** collection.
 

 



```vb
Dim objWebNavBarSet As WebNavigationBarSet 
For Each objWebNavBarSet In ActiveDocument.WebNavigationBarSets 
 objWebNavBarSet.DeleteSetAndInstances 
Next objWebNavBarSet
```

There are three properties that concern horizontally oriented Web navigation bars. Use  **WebNavigationBarSet**. **IsHorizontal** to determine the orientation of the navigation bar set. The **ChangeOrientation** method is used to set the orientation of the Web navigation bar set. If the orientation is set to **horizontal**, **HorizontalAlignment** and **HorizontalButtonCount** properties can then be set. The following example adds the first navigation bar in the **WebNavigationBarSets** collection of the active document to each page that has the **AddHyperlinkToWebNavbar** property set to **True** or the **Page.WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True**, and then sets the button style to **small**. A test is performed to determine whether the navigation bar set is horizontal or not. If it is not, the **ChangeOrientation** method is called and the orientation is set to **horizontal**. After the navigation bar is oriented horizontally, the horizontal button count is set to **3** and the horizontal alignment of the buttons is set to **left**.
 

 



```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignLeft 
End With
```


## Methods



|Name|
|:-----|
|[AddToEveryPage](Publisher.WebNavigationBarSet.AddToEveryPage.md)|
|[ChangeOrientation](Publisher.WebNavigationBarSet.ChangeOrientation.md)|
|[DeleteSetAndInstances](Publisher.WebNavigationBarSet.DeleteSetAndInstances.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.WebNavigationBarSet.Application.md)|
|[AutoUpdate](Publisher.WebNavigationBarSet.AutoUpdate.md)|
|[ButtonStyle](Publisher.WebNavigationBarSet.ButtonStyle.md)|
|[Design](Publisher.WebNavigationBarSet.Design.md)|
|[HorizontalAlignment](Publisher.WebNavigationBarSet.HorizontalAlignment.md)|
|[HorizontalButtonCount](Publisher.WebNavigationBarSet.HorizontalButtonCount.md)|
|[IsHorizontal](Publisher.WebNavigationBarSet.IsHorizontal.md)|
|[Links](Publisher.WebNavigationBarSet.Links.md)|
|[Name](Publisher.WebNavigationBarSet.Name.md)|
|[Parent](Publisher.WebNavigationBarSet.Parent.md)|
|[ShowSelected](Publisher.WebNavigationBarSet.ShowSelected.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]