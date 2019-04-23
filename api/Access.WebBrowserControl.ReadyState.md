---
title: WebBrowserControl.ReadyState property (Access)
keywords: vbaac10.chm14362
f1_keywords:
- vbaac10.chm14362
ms.prod: access
api_name:
- Access.WebBrowserControl.ReadyState
ms.assetid: 49ba1888-9a1e-ea35-18ed-b3bfbbfd3f31
ms.date: 03/26/2019
localization_priority: Normal
---


# WebBrowserControl.ReadyState property (Access)

Gets the status of the specified web browser control. Read-only **[AcWebBrowserState](Access.AcWebBrowserState.md)**.


## Syntax

_expression_.**ReadyState**

_expression_ A variable that represents a **[WebBrowserControl](Access.WebBrowserControl.md)** object.


## Remarks

When the web browser control does not contain a document, the value of this property is **acUninitialized**. Other values indicate the control state when it navigates to a new document.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]