---
title: EdgeBrowserControl.TrustedDomains property (Access)
keywords: vbaac10.chm14367,vbaac10.chm5911
f1_keywords:
- vbaac10.chm14367,vbaac10.chm5911
ms.prod: access
api_name:
- Access.EdgeBrowserControl.TrustedDomains
ms.assetid: 1c444ac8-d021-49ef-8687-f05f21e10bdb
ms.date: 03/08/2023
ms.localizationpriority: medium
---


# EdgeBrowserControl.TrustedDomains property (Access)

Read/write. Allows you to specify a **table name** who's first column contains domains the browser is allowed to renavigate to


## Syntax

_expression_.**TrustedDomains**

_expression_ A variable that represents an **[EdgeBrowserControl](Access.EdgeBrowserControl.md)** object.

## Remarks
If the **Control Source** is bound to a field, the browser will already be allowed to navigate to those values. This property is useful if you want to allow redirects that happen during logins, or to allow links to other domains on the starting web page to work.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]