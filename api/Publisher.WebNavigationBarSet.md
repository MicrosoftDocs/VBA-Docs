---
title: WebNavigationBarSet object (Publisher)
keywords: vbapb10.chm8585215
f1_keywords:
- vbapb10.chm8585215
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet
ms.assetid: 03b31cc1-5b24-1a16-710c-73755298066e
ms.date: 06/04/2019
localization_priority: Normal
---


# WebNavigationBarSet object (Publisher)

Represents a web navigation bar set for the current document. The **WebNavigationBarSet** object is a member of the **[WebNavigationBarSets](publisher.webnavigationbarsets.md)** collection, which includes all the web navigation bar sets in the current document.
 
## Remarks

To add the specified web navigation bar to every page of a document, use the _Left_, _Top_, and _Width_ parameters of the **AddToEveryPage** method, where _Left_ is the position of the left edge of the shape, _Top_ is the position of the top edge of the shape, and _Width_ is the width of the shape representing the web navigation bar set. 

To remove the web navigation bar set and every instance of it from a document, use the **DeleteSetAndInstances** method. 

The following concern horizontally oriented web navigation bars:

- Use the **IsHorizontal** property to determine the orientation of the navigation bar set. 
- Use the **ChangeOrientation** method to set the orientation of the web navigation bar set. 
- If the orientation is set to **horizontal**, you can then set the **HorizontalAlignment** and **HorizontalButtonCount** properties. 

## Example

The following example adds the first web navigation bar set to every page that has the **AddToEveryPage** method set to **True** when adding the page, or the **[WebPageOptions.IncludePageOnNewWebNavigationBars](publisher.webpageoptions.includepageonnewwebnavigationbars.md)** property set to **True**.

```vb
Dim objWebNavBarSet as WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets(1) 
objWebNavBarSet.AddToEveryPage Left:=50, Top:=10, Width:=500
```

<br/>

The following example deletes all instances of each **WebNavigationBarSet** object in the **WebNavigationBarSets** collection.

```vb
Dim objWebNavBarSet As WebNavigationBarSet 
For Each objWebNavBarSet In ActiveDocument.WebNavigationBarSets 
 objWebNavBarSet.DeleteSetAndInstances 
Next objWebNavBarSet
```

<br/>

The following example adds the first navigation bar in the **WebNavigationBarSets** collection of the active document to each page that has the **AddToEveryPage** method set to **True**, or the **IncludePageOnNewWebNavigationBars** property set to **True**, and then sets the button style to **small**. A test is performed to determine whether the navigation bar set is horizontal. If it is not, the **ChangeOrientation** method is called and the orientation is set to **horizontal**. After the navigation bar is oriented horizontally, the horizontal button count is set to **3** and the horizontal alignment of the buttons is set to **left**.

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

- [AddToEveryPage](Publisher.WebNavigationBarSet.AddToEveryPage.md)
- [ChangeOrientation](Publisher.WebNavigationBarSet.ChangeOrientation.md)
- [DeleteSetAndInstances](Publisher.WebNavigationBarSet.DeleteSetAndInstances.md)

## Properties

- [Application](Publisher.WebNavigationBarSet.Application.md)
- [AutoUpdate](Publisher.WebNavigationBarSet.AutoUpdate.md)
- [ButtonStyle](Publisher.WebNavigationBarSet.ButtonStyle.md)
- [Design](Publisher.WebNavigationBarSet.Design.md)
- [HorizontalAlignment](Publisher.WebNavigationBarSet.HorizontalAlignment.md)
- [HorizontalButtonCount](Publisher.WebNavigationBarSet.HorizontalButtonCount.md)
- [IsHorizontal](Publisher.WebNavigationBarSet.IsHorizontal.md)
- [Links](Publisher.WebNavigationBarSet.Links.md)
- [Name](Publisher.WebNavigationBarSet.Name.md)
- [Parent](Publisher.WebNavigationBarSet.Parent.md)
- [ShowSelected](Publisher.WebNavigationBarSet.ShowSelected.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]