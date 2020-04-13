---
title: SolutionsModule.Session property (Outlook)
keywords: vbaol11.chm3362
f1_keywords:
- vbaol11.chm3362
ms.prod: outlook
api_name:
- Outlook.SolutionsModule.Session
ms.assetid: 28a67ff1-1427-2852-cf00-1aeb926ba8dc
ms.date: 06/08/2017
localization_priority: Normal
---


# SolutionsModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a '[SolutionsModule](Outlook.SolutionsModule.md)' object.


## Remarks

Returns  **Null** (**Nothing** in Visual Basic) if there is no logged-on session.

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function.




```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```




```vb
Set objSession = Application.Session
```


## See also


[SolutionsModule Object](Outlook.SolutionsModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]