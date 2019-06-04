---
title: Application.Language property (Publisher)
keywords: vbapb10.chm131091
f1_keywords:
- vbapb10.chm131091
ms.prod: publisher
api_name:
- Publisher.Application.Language
ms.assetid: 2fcfbec9-0c84-43d5-8c53-5b73bca17e3d
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.Language property (Publisher)

Returns a **Long** that represents the language selected for the Microsoft Publisher user interface. Read-only.


## Syntax

_expression_.**Language**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

Long


## Remarks

The **LanguageID** property value can be one of the **[MsoLanguageID](Office.MsoLanguageID.md)** constants declared in the Microsoft Office type library.


## Example

This example displays a message stating whether the language selected for the Publisher user interface is U.S. English.

```vb
Sub LangSetting() 
 If Application.Language = msoLanguageIDEnglishUS Then 
 MsgBox "The user interface language is U.S. English." 
 Else 
 MsgBox "The user interface language is not U.S. English." 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]