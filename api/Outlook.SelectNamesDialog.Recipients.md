---
title: SelectNamesDialog.Recipients property (Outlook)
keywords: vbaol11.chm827
f1_keywords:
- vbaol11.chm827
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.Recipients
ms.assetid: 8b939af1-b266-55ad-f9ad-8802ac2e0930
ms.date: 06/08/2017
localization_priority: Normal
---


# SelectNamesDialog.Recipients property (Outlook)

Returns a **[Recipients](Outlook.Recipients.md)** collection object that represents the recipients selected in the **Select Names** dialog, or sets a **Recipients** collection object that represents the initial recipients to be displayed in the **Select Names** dialog box. Read/write.


## Syntax

_expression_. `Recipients`

_expression_ A variable that represents a [SelectNamesDialog](Outlook.SelectNamesDialog.md) object.


## Remarks

This property specifies a **Recipients** collection object that has a **[Recipients.Count](Outlook.Recipients.Count.md)** equal to the total number of recipients in the **To**,  **Cc**, and  **Bcc** edit boxes.

If you do not set this property before displaying the  **Select Names** dialog box, then the **Recipients** object represented by **SelectNamesDialog.Recipients** will have a **Recipients.Count** equal to zero.

If the user does not select any names from the  **Select Names** dialog box and clicks **OK**,  **SelectNamesDialog.Recipients** will return a **Recipients** collection object with **Recipients.Count** equal to zero.


## See also


[SelectNamesDialog Object](Outlook.SelectNamesDialog.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]