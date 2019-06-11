---
title: Page.IsTrailing property (Publisher)
keywords: vbapb10.chm131101
f1_keywords:
- vbapb10.chm131101
ms.prod: publisher
api_name:
- Publisher.Page.IsTrailing
ms.assetid: e0ed15dc-d2e8-d6b7-913d-4e72b2817e88
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.IsTrailing property (Publisher)

**True** if the specified **Page** object is a trailing page of a two-page spread. Read-only **Boolean**.


## Syntax

_expression_.**IsTrailing**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

Boolean


## Example

The following example displays for each page whether the page is a trailing or leading page in the publication.

```vb
Dim objPage As Page 
Dim strPageInfo As String 
For Each objPage In ActiveDocument.Pages 
 strPageInfo = "Page number " & objPage.PageNumber 
 If objPage.IsLeading Then 
 strPageInfo = strPageInfo & " is a leading page." & Chr(13) 
 ElseIf objPage.IsTrailing Then 
 strPageInfo = strPageInfo & " is a trailing page." & Chr(13) 
 End If 
 MsgBox strPageInfo 
Next objPage
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]