---
title: WebNavigationBarSets object (Publisher)
keywords: vbapb10.chm8519679
f1_keywords:
- vbapb10.chm8519679
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSets
ms.assetid: 0c4f62c7-b7b2-a7bc-60f8-8097fe99fe58
ms.date: 06/04/2019
localization_priority: Normal
---


# WebNavigationBarSets object (Publisher)

A collection of all the **[WebNavigationBarSet](publisher.webnavigationbarset.md)** objects in the current document. Each **WebNavigationBarSet** object represents a web navigation bar set consisting of hyperlinks.
 

## Remarks

By default there are two **WebNavigationBarSet** objects on each web wizard page; one is text-only and the other is vertical. These objects correspond to the design of the wizard regardless of whether a navigation bar is used on the page.
 
Use the **[WebNavigationBarSets](publisher.document.webnavigationbarsets.md)** property of the current document to return a **WebNavigationBarSets** object. 
 
Use **Item** (_index_), where _index_ is the index number, to return a **WebNavigationBarSet** object from the collection. 

Use the **Count** property to return the number of web navigation bar sets in the collection. 

To add the specified web navigation bar to every page of a document, use the _Left_, _Top_, and _Width_ parameters of the **AddToEveryPage** method, where _Left_ is the distance from the left of the page to the left edge of the navigation bar, _Top_ is the distance from the top of the page to the top edge of the navigation bar, and _Width_ is the width of the navigation bar. 
 

## Example

The following example sets an object variable to the **WebNavigationBarSets** collection of the active document.

```vb
Dim objWebNavBarSets As WebNavigationBarSets 
Set objWebNavBarSets = ActiveDocument.WebNavigationBarSets
```

<br/>

The following example returns the first web navigation bar set from the **WebNavigationBarSets** collection.

```vb
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.Item(1)
```

<br/>

The previous example can also be accomplished by using **WebNavigationBarSets** (_index_), where _index_ is the index number, to return a **WebNavigationBarSet** object.
 
```vb
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets(1)
```

<br/>

The previous example can also be accomplished by using **WebNavigationBarSets** (_index_), where _index_ is a string indicating the name of the web navigation bar set to return.
 
```vb
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets("WebNavBarSet1")
```

<br/>

This example displays the number of web navigation bar sets in the current document.

```vb
MsgBox ActiveDocument.WebNavigationBarSets.Count 

```

<br/>

The following example adds the navigation bar named WebNavBarSet1 to every page in the current publication.

```vb
ActiveDocument.WebNavigationBarSets.Item _ 
 ("WebNavBarSet1").AddToEveryPage _ 
 Left:=50, Top:=25
```


## Methods

- [AddSet](Publisher.WebNavigationBarSets.AddSet.md)
- [Item](Publisher.WebNavigationBarSets.Item.md)

## Properties

- [Application](Publisher.WebNavigationBarSets.Application.md)
- [Count](Publisher.WebNavigationBarSets.Count.md)
- [Parent](Publisher.WebNavigationBarSets.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]