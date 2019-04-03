---
title: Tab.Index Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 2cacd35e-edd4-6733-e932-a05114134754
ms.date: 06/08/2017
localization_priority: Normal
---


# Tab.Index Property (Outlook Forms Script)

Returns or sets an  **Integer** that specifies the position of a **[Tab](Outlook.tab.md)** object within a **[Tabs](Outlook.tabs.md)** collection. Read/write.


## Syntax

_expression_.**Index**

_expression_ A variable that represents a  **Tab** object.


## Remarks

The  **Index** property specifies the order in which tabs appear. Changing the value of **Index** visually changes the order of tabs on a **[TabStrip](Outlook.tabstrip.md)**. The index value for the first tab is zero, the index value of the second tab is one, and so on.

In a  **MultiPage**,  **Index** refers to a **Page** as well as the page's **Tab**. In a  **TabStrip**,  **Index** refers to the tab only.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]