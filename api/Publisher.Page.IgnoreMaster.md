---
title: Page.IgnoreMaster property (Publisher)
keywords: vbapb10.chm393233
f1_keywords:
- vbapb10.chm393233
ms.prod: publisher
api_name:
- Publisher.Page.IgnoreMaster
ms.assetid: 53cd7b4b-4164-c6d3-766f-885a056d9b2b
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.IgnoreMaster property (Publisher)

**True** for Microsoft Publisher to ignore the master page formatting for the specified page. Read/write **Boolean**.


## Syntax

_expression_.**IgnoreMaster**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

Boolean


## Example

This example adds a red star in the upper-left corner of the master page so that it shows up on each page; it then adds a few new pages and sets one of the pages to ignore the master page so that the shape doesn't show on it.

```vb
Sub AddNewPageIgnoreMaster() 
 Dim pgNew As Page 
 
 With ActiveDocument 
 .MasterPages(1).Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=50, Top:=50, Width:=50, Height:=50).Fill.ForeColor _ 
 .CMYK.SetCMYK Cyan:=0, Magenta:=255, Yellow:=255, Black:=0 
 .Pages.Add Count:=1, After:=1 
 Set pgNew = .Pages.Add(Count:=1, After:=1) 
 pgNew.IgnoreMaster = True 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]