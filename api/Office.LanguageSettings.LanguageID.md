---
title: LanguageSettings.LanguageID property (Office)
keywords: vbaof11.chm231001
f1_keywords:
- vbaof11.chm231001
ms.prod: office
api_name:
- Office.LanguageSettings.LanguageID
ms.assetid: a1efbab6-000f-d87e-296b-b58be9ad5194
ms.date: 01/18/2019
localization_priority: Normal
---


# LanguageSettings.LanguageID property (Office)

Gets an **[MsoAppLanguageID](office.msoapplanguageid.md)** constant representing the locale identifier (LCID) for the install language, the user interface language, or the Help language. Read-only.


## Syntax

_expression_.**LanguageID**(_Id_)

_expression_ A variable that represents a **[LanguageSettings](Office.LanguageSettings.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Id_|Required|**MsoAppLanguageID**|Returns one of the **MsoAppLanguageID** enumerations.|

## Example

This Microsoft Excel example checks the **LanguageID** property settings for the user interface and execution mode to verify that they are set to the same LCID. The example returns an error if there is a discrepancy.


```vb
If Application.LanguageSettings.LanguageID(msoLanguageIDExeMode) _ 
 > Application.LanguageSettings.LanguageID(msoLanguageIDUI) _ 
 Then MsgBox "The user interface language and execution " & _ 
 "mode are different."
```


## See also

- [LanguageSettings object members](overview/Library-Reference/languagesettings-members-office.md)





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]