---
title: Options.PathForPublications property (Publisher)
keywords: vbapb10.chm1048597
f1_keywords:
- vbapb10.chm1048597
ms.prod: publisher
api_name:
- Publisher.Options.PathForPublications
ms.assetid: d33d5eab-eb52-b533-8968-31ddb5e12d99
ms.date: 06/11/2019
localization_priority: Normal
---


# Options.PathForPublications property (Publisher)

Returns a **String** that represents the default folder for publications. Read-only.


## Syntax

_expression_.**PathForPublications**

_expression_ A variable that represents an **[Options](Publisher.Options.md)** object.


## Return value

String


## Example

This example returns the current default path for publications (corresponds to the default path setting on the **General** tab in the **Options** dialog box, **Tools** menu).

```vb
Sub PubPath() 
 Dim strPubPath 
 strPubPath = Options.PathForPublications 
 MsgBox strPubPath 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]