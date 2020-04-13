---
title: Application.DisplayScreenTips property (Word)
keywords: vbawd10.chm158335075
f1_keywords:
- vbawd10.chm158335075
ms.prod: word
api_name:
- Word.Application.DisplayScreenTips
ms.assetid: 07a03053-4973-27e2-6f0c-f67ff03c8bcf
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DisplayScreenTips property (Word)

 **True** if comments, footnotes, endnotes, and hyperlinks are displayed as tips. Text marked as having comments is highlighted. Read/write **Boolean**.


## Syntax

_expression_. `DisplayScreenTips`

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example enables Word to display comments, footnotes, and endnotes as tips. Also, text marked as having comments is highlighted.


```vb
Application.DisplayScreenTips = True
```

This example returns the current status of the **ScreenTips** checkbox in the **Show** area on the **View** tab in the **Options** dialog box.




```vb
temp = Application.DisplayScreenTips
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]