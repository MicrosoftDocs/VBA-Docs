---
title: Page.Visible Property (Outlook Forms Script)
keywords: olfm10.chm2002200
f1_keywords:
- olfm10.chm2002200
ms.prod: outlook
ms.assetid: 2023a10d-72d3-893a-9044-9f39f6cd0539
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Visible Property (Outlook Forms Script)

Returns or sets a **Boolean** that specifies whether a **[Page](Outlook.page.md)** is visible or hidden. Read/write.


## Syntax

_expression_.**Visible**

_expression_ A variable that represents a **Page** object.


## Remarks

 **True** to specify the page is visible (default), **False** to specify the page is hidden.

Use the  **Visible** property to control access to information without displaying it. For example, you could use the value of a control on a hidden form as the criteria for a query.

All pages are visible at design time.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]