---
title: Window.DisplayScreenTips property (Word)
keywords: vbawd10.chm157417494
f1_keywords:
- vbawd10.chm157417494
ms.prod: word
api_name:
- Word.Window.DisplayScreenTips
ms.assetid: fc90fe70-ed5d-b02c-63fd-59696ed70465
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.DisplayScreenTips property (Word)

 **True** if comments, footnotes, endnotes, and hyperlinks are displayed as tips. Read/write **Boolean**.


## Syntax

_expression_. `DisplayScreenTips`

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Remarks

Text marked as having comments is highlighted.


## Example

This example enables Word to display comments, footnotes, and endnotes as tips. Also, text marked as having comments is highlighted.


```vb
Application.DisplayScreenTips = True
```

This example returns the current status of the  **Show document tooltips on hover** checkbox in the **Page display options** section on the **Display** tab of the **Word Options** dialog box.




```vb
temp = Application.DisplayScreenTips
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]