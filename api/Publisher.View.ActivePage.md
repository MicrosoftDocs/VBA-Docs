---
title: View.ActivePage property (Publisher)
keywords: vbapb10.chm327683
f1_keywords:
- vbapb10.chm327683
ms.prod: publisher
api_name:
- Publisher.View.ActivePage
ms.assetid: 29289fb2-6692-4cb5-a9e2-b2edb9e9cd7e
ms.date: 06/15/2019
localization_priority: Normal
---


# View.ActivePage property (Publisher)

Returns a **[Page](Publisher.Page.md)** object that represents the page currently displayed in the Microsoft Publisher window.


## Syntax

_expression_.**ActivePage**

_expression_ A variable that represents a **[View](Publisher.View.md)** object.


## Return value

Page


## Example

This example saves the active page as a JPEG picture. Note that `PathToFile` must be replaced with a valid file path for this example to work.

```vb
Sub SavePageAsPicture() 
 ActiveView.ActivePage.SaveAsPicture _ 
 FileName:="PathToFile" 
End Sub
```

<br/>

This example adds a horizontal ruler guide and a vertical ruler guide to the active page that intersect at the center point of the page.

```vb
Sub SetRulerGuidesOnActivePage() 
 Dim intHeight As Integer 
 Dim intWidth As Integer 
 
 With ActiveView.ActivePage 
 intHeight = .Height / 2 
 intWidth = .Width / 2 
 With .RulerGuides 
 .Add Position:=intHeight, Type:=pbRulerGuideTypeHorizontal 
 .Add Position:=intWidth, Type:=pbRulerGuideTypeVertical 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]