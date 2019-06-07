---
title: LayoutGuides.MirrorGuides property (Publisher)
keywords: vbapb10.chm1114119
f1_keywords:
- vbapb10.chm1114119
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.MirrorGuides
ms.assetid: 8e6ff709-21e0-2286-5d75-c7ebea05fd26
ms.date: 06/08/2019
localization_priority: Normal
---


# LayoutGuides.MirrorGuides property (Publisher)

Returns or sets a **Boolean** indicating whether Microsoft Publisher creates mirror guide positions for a book fold publication. 

**True** if Publisher creates mirror guide positions for separate left and right pages in a book fold publication; **False** if the same margin, row, and column guides are applied to all pages in the publication. Read/write.


## Syntax

_expression_.**MirrorGuides**

_expression_ A variable that represents a **[LayoutGuides](Publisher.LayoutGuides.md)** object.


## Return value

Boolean


## Remarks

When the **MirrorGuides** property is **True**, margin settings apply to right-facing pages and are mirrored for left-facing pages. 

In addition, when set to **True**, the **MirrorGuides** property sets the publication to use two master pages instead of the default of one. The first master page is for all left-facing pages, and the second is for all right-facing pages in the publication. 

For more information, see the **[MasterPages](Publisher.MasterPages.md)** object.


## Example

The following example sets Publisher to create mirror guides for a book fold publication and sets the inside and outside margins of each two-page spread. The example sets the left and right margin values for right-facing pages, and Publisher mirrors these values for left-facing pages.

```vb
With ActiveDocument.LayoutGuides 
 .MirrorGuides = True 
 .MarginLeft = 48 
 .MarginRight = 96 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]