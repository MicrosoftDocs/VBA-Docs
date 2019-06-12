---
title: Section.ContinueNumbersFromPreviousSection property (Publisher)
keywords: vbapb10.chm7405575
f1_keywords:
- vbapb10.chm7405575
ms.prod: publisher
api_name:
- Publisher.Section.ContinueNumbersFromPreviousSection
ms.assetid: a3d64f14-dc65-4fb1-5079-0fdf2e3f8f38
ms.date: 06/13/2019
localization_priority: Normal
---


# Section.ContinueNumbersFromPreviousSection property (Publisher)

**True** if the specified section continues the numbering from the previous section. Read/write **Boolean**.


## Syntax

_expression_.**ContinueNumbersFromPreviousSection**

_expression_ A variable that represents a **[Section](Publisher.Section.md)** object.


## Return value

Boolean


## Example

The following example adds three pages to the publication, adds a new section after the first page, and then sets the **ContinueNumbersFromPreviousSection** to **False** for the new section.

```vb
Dim objSection As Section 
ActiveDocument.Pages.Add Count:=3, After:=1 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2) 
objSection.ContinueNumbersFromPreviousSection = False 
 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]