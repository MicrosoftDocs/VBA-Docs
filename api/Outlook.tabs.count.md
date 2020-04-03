---
title: Tabs.Count Property (Outlook Forms Script)
keywords: olfm10.chm2001000
f1_keywords:
- olfm10.chm2001000
ms.prod: outlook
ms.assetid: 1424d686-d082-26f8-8312-942aad178813
ms.date: 06/08/2017
localization_priority: Normal
---


# Tabs.Count Property (Outlook Forms Script)

Returns a **Long** that represents the number of objects in a collection. Read-only.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **Tabs** object.


## Remarks

Note that the index value for the first tab of a collection is zero, the value for the second tab is one, and so on. For example, if a **[TabStrip](Outlook.tabstrip.md)** contains two tabs, the indexes of the tabs are 0 and 1, and the value of **Count** is 2.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]