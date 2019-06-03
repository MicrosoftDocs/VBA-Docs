---
title: WebNavigationBarHyperlinks object (Publisher)
keywords: vbapb10.chm540671
f1_keywords:
- vbapb10.chm540671
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarHyperlinks
ms.assetid: 4dfa7273-4770-d77c-275c-6b7eeae04aa5
ms.date: 06/04/2019
localization_priority: Normal
---


# WebNavigationBarHyperlinks object (Publisher)

Represents a collection of all the **[Hyperlink](publisher.hyperlink.md)** objects of the specified **[WebNavigationBarSet](publisher.webnavigationbarset.md)** object.
 
## Remarks

Use the **[Links](publisher.webnavigationbarset.links.md)** property of the **WebNavigationBarSet** object to return a **WebNavigationBarHyperlinks** object. 

Use the **Count** property to return a **Long** representing the number of hyperlinks in the **WebNavigationBarHyperlinks** collection of the specified **WebNavigationBarSet** object. 

Use **Item** (_index_), where _index_ is the index number, to return a specific **Hyperlink** object from the collection. 

## Example

The following example adds a hyperlink to the first **WebNavigationBarSet** of the active document.

```vb
Dim objWebNavLinks As WebNavigationBarHyperlinks 
Set objWebNavLinks = ActiveDocument.WebNavigationBarSets(1).Links 
objWebNavLinks.Add Address:="www.microsoft.com", _ 
 TextToDisplay:="Microsoft"
```

<br/>

The following example displays the number of hyperlinks in the first **WebNavigationBarSet** of the active document.

```vb
MsgBox ActiveDocument.WebNavigationBarSets(1).Links.Count
```

<br/>

This example displays the displayed text of the first item in the **WebNavigationBarHyperlinks** collection of the first **WebNavigationBarSet** of the active document.

```vb
MsgBox ActiveDocument.WebNavigationBarSets(1).Links.Item(1).TextToDisplay
```


## Methods

- [Add](Publisher.WebNavigationBarHyperlinks.Add.md)
- [Item](Publisher.WebNavigationBarHyperlinks.Item.md)

## Properties

- [Application](Publisher.WebNavigationBarHyperlinks.Application.md)
- [Count](Publisher.WebNavigationBarHyperlinks.Count.md)
- [Parent](Publisher.WebNavigationBarHyperlinks.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]