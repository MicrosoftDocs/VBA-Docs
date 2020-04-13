---
title: Page.Index Property (Outlook Forms Script)
keywords: olfm10.chm2001280
f1_keywords:
- olfm10.chm2001280
ms.prod: outlook
ms.assetid: 91e67439-ea23-9ac8-6065-31af7be0b303
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Index Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies the position of a **[Page](Outlook.page.md)** object in a **[Pages](Outlook.pages(object).md)** collection. Read/write.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a **Page** object.


## Remarks

The **Index** property specifies the order in which tabs appear. Changing the value of **Index** visually changes the order of pages in a **[MultiPage](Outlook.multipage.md)**. The index value for the first page is zero, the index value of the second page is one, and so on.

In a **MultiPage**,  **Index** refers to a **Page** as well as the page's **[Tab](Outlook.tab.md)**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]