---
title: Section.ShowHeaderFooterOnFirstPage property (Publisher)
keywords: vbapb10.chm7405574
f1_keywords:
- vbapb10.chm7405574
ms.prod: publisher
api_name:
- Publisher.Section.ShowHeaderFooterOnFirstPage
ms.assetid: 6c814884-9bee-72ae-3a40-5118bebd6f02
ms.date: 06/13/2019
localization_priority: Normal
---


# Section.ShowHeaderFooterOnFirstPage property (Publisher)

**True** if the header and footer of the specified section will be visible. Read/write **Boolean**.


## Syntax

_expression_.**ShowHeaderFooterOnFirstPage**

_expression_ A variable that represents a **[Section](Publisher.Section.md)** object.


## Return value

Boolean


## Example

The following example adds a new section starting on the second page of the active document, adds header and footer text to the master page, and then sets the **ShowHeaderFooterOnFirstPage** property to **True**.

```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2) 
With ActiveDocument.Pages(2).Master 
 .Header.TextRange.Text = "Page " & .PageNumber & " header." 
 .Footer.TextRange.Text = "Page " & .PageNumber & " footer." 
End With 
objSection.ShowHeaderFooterOnFirstPage = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]