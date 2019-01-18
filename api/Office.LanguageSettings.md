---
title: LanguageSettings object (Office)
keywords: vbaof11.chm231000
f1_keywords:
- vbaof11.chm231000
ms.prod: office
api_name:
- Office.LanguageSettings
ms.assetid: 936f7d61-87e5-e153-08d4-f8c5c8ef0710
ms.date: 06/08/2017
localization_priority: Normal
---


# LanguageSettings object (Office)

Returns information about the language settings in a Microsoft Office application.


## Remarks

Use Application.LanguageSettings.LanguageID( _MsoAppLanguageID_ ), where[MsoAppLanguageID](Office.MsoAppLanguageID.md) is a constant used to return locale identifier (LCID) information to the specified application.


## Example

The following example returns the install language, user interface language, and Help language LCIDs in a message box.


```vb
MsgBox "The following locale IDs are registered " &amp; _ 
 "for this application: Install Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDInstall) &amp; _ 
 " User Interface Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDUI) &amp; _ 
 " Help Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDHelp)
```

Use  **Application.LanguageSettings.LanguagePreferredForEditing** to determine which LCIDs are registered as preferred editing languages for the application, as in the following example.




```vb
If Application.LanguageSettings. _ 
 LanguagePreferredForEditing(msoLanguageIDEnglishUS) Then 
 MsgBox "U.S. English is one of the chosen editing languages." 
End If
```


## See also


[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[LanguageSettings Object Members](./overview/Library-Reference/languagesettings-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]