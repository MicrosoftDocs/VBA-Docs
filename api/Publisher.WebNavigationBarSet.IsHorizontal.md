---
title: WebNavigationBarSet.IsHorizontal property (Publisher)
keywords: vbapb10.chm8519686
f1_keywords:
- vbapb10.chm8519686
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.IsHorizontal
ms.assetid: d3bbb0b0-8d06-7d46-1ef7-fef0a3e846b7
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarSet.IsHorizontal property (Publisher)

**True** if the specified **WebNavigationBarSet** has a horizontal orientation. Read-only **Boolean**.


## Syntax

_expression_.**IsHorizontal**

_expression_ A variable that represents a **[WebNavigationBarSet](Publisher.WebNavigationBarSet.md)** object.


## Return value

Boolean


## Remarks

This property is used to determine the orientation of the navigation bar set prior to setting certain properties that assume a horizontal orientation such as the **HorizontalAlignment** property.


## Example

This example adds the first navigation bar in the **[WebNavigationBarSets](Publisher.WebNavigationBarSets.md)** collection of the active document to each page, and then sets the button style to small. 

A test is performed to determine whether the navigation bar set is horizontal. If it is not, the **[ChangeOrientation](Publisher.WebNavigationBarSet.ChangeOrientation.md)** method is called and the orientation is set to horizontal. After the navigation bar is oriented horizontally, the horizontal button count is set to 3 and the horizontal alignment of the buttons is set to **pbnbAlignLeft**.


```vb
Dim objWebNav As WebNavigationBarSet 
Set objWebNav = ActiveDocument.WebNavigationBarSets(1) 
With objWebNav 
 .AddToEveryPage Left:=10, Top:=10 
 .ButtonStyle = pbnbButtonStyleSmall 
 If .IsHorizontal = False Then 
 .ChangeOrientation pbNavBarOrientHorizontal 
 End If 
 .HorizontalButtonCount = 3 
 .HorizontalAlignment = pbnbAlignLeft 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]