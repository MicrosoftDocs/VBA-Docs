---
title: Page.ReaderSpread property (Publisher)
keywords: vbapb10.chm393238
f1_keywords:
- vbapb10.chm393238
ms.prod: publisher
api_name:
- Publisher.Page.ReaderSpread
ms.assetid: 32823d2d-4bcd-a5a6-1ad1-ca1035d4fdea
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.ReaderSpread property (Publisher)

Returns a **[ReaderSpread](Publisher.ReaderSpread.md)** object that represents the reader spread of the specified page.


## Syntax

_expression_.**ReaderSpread**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

ReaderSpread


## Example

This example checks to see if the reader spread for the specified page includes fewer than two pages. If it does, it changes the reader spread to include two pages.

```vb
Sub SetFacingPages() 
 With ActiveDocument.Pages(2).ReaderSpread 
 If .PageCount < 2 Then _ 
 ActiveDocument.ViewTwoPageSpread = True 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]