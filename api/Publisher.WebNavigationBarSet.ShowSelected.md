---
title: WebNavigationBarSet.ShowSelected property (Publisher)
keywords: vbapb10.chm8519696
f1_keywords:
- vbapb10.chm8519696
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.ShowSelected
ms.assetid: c8229f03-a043-a280-84f9-f75a430c3903
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.ShowSelected property (Publisher)

**True** if the selected button is highlighted for the specified **WebNavigationBarSet** object. Read/write **Boolean**.


## Syntax

_expression_.**ShowSelected**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

Boolean


## Example

The following example adds a new web navigation bar to the active document, adds it to every page, and then sets the **ShowSelected** property to **False** so that the selected button is not highlighted in the navigation bar.

```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets.AddSet(Name:="newNavBar") 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleSmall 
 .ShowSelected = False 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]