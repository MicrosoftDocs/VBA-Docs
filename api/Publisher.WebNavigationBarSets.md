---
title: WebNavigationBarSets Object (Publisher)
keywords: vbapb10.chm8519679
f1_keywords:
- vbapb10.chm8519679
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSets
ms.assetid: 0c4f62c7-b7b2-a7bc-60f8-8097fe99fe58
ms.date: 06/08/2017
localization_priority: Normal
---


# WebNavigationBarSets Object (Publisher)

A collection of all the  **WebNavigationBarSet** objects in the current document. Each **WebNavigationBarSet** represents a Web navigation bar set consisting of hyperlinks.
 


## Remarks

By default there are two  **WebNavigationBarSet** objects on each Web wizard page; one is text-only and the other is vertical. These objects correspond to the design of the wizard regardless of whether or not a navigation bar is used on the page.
 

 

## Example

Use the  **WebNavigationBarSets** property of the current document to return a **WebNavigationBarSet** object. The following example sets an object variable to the **WebNavigationBarSets** collection of the active document.
 

 

```vb
Dim objWebNavBarSets As WebNavigationBarSets 
Set objWebNavBarSets = ActiveDocument.WebNavigationBarSets
```

Use  **WebNavigationBarSets** **.Item** (index), where index is the index number, to return a **WebNavigationBarSet** object from the collection. The following example returns the first Web navigation bar set from the **WebNavigationBarSets** collection.
 

 



```vb
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.Item(1)
```

The previous example can also be accomplished using  **WebNavigationBarSets** (index), where index is the index number, to return a **WebNavigationBarSet** object.
 

 



```vb
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets(1)
```

The previous example can also be accomplished using  **WebNavigationBarSets** (index) where index is a string indicating the name of the Web navigation bar set to return.
 

 



```vb
Dim objWebNavBarSet As WebNavigationBarSet 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets("WebNavBarSet1")
```

Use  **WebNavigationBarSets** **.Count** to return the number of Web navigation bar sets in the collection. This example displays the number of Web navigation bar sets in the current document.
 

 



```vb
MsgBox ActiveDocument.WebNavigationBarSets.Count 

```

Use  **WebNavigationBarSets** **.AddToEveryPage** (Left, Top, [Width]), where Left is the distance from the left of the page to the left edge of the navigation bar, Top is the distance form the top of the page to the top edge of the navigation bar, and Width is the width of the navigaion bar, to add the specified navigation bar to every page. The following example adds the navigation bar named "WebNavBar1" to every page in the current publication.
 

 



```vb
ActiveDocument.WebNavigationBarSets.Item _ 
 ("WebNavBarSet1").AddToEveryPage _ 
 Left:=50, Top:=25
```


## Methods



|Name|
|:-----|
|[AddSet](Publisher.WebNavigationBarSets.AddSet.md)|
|[Item](Publisher.WebNavigationBarSets.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.WebNavigationBarSets.Application.md)|
|[Count](Publisher.WebNavigationBarSets.Count.md)|
|[Parent](Publisher.WebNavigationBarSets.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]