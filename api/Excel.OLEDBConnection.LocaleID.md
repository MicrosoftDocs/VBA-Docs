---
title: OLEDBConnection.LocaleID property (Excel)
keywords: vbaxl10.chm794107
f1_keywords:
- vbaxl10.chm794107
ms.prod: excel
api_name:
- Excel.OLEDBConnection.LocaleID
ms.assetid: 6a92f9ca-247a-8da8-a32e-ec239380894a
ms.date: 05/02/2019
localization_priority: Normal
---


# OLEDBConnection.LocaleID property (Excel)

Returns or sets the locale identifier for the specified connection. Read/write.


## Syntax

_expression_.**LocaleID**

_expression_ A variable that represents an **[OLEDBConnection](Excel.OLEDBConnection.md)** object.


## Return value

**Integer**


## Remarks

Before you set the **LocaleID** property to a new locale, you must set the **[RetrieveInOfficeUILang](Excel.OLEDBConnection.RetrieveInOfficeUILang.md)** property to **False**. For more information about valid Locale ID (LCID) values, see the [LCID-Locale Mapping Table](https://docs.microsoft.com/openspecs/windows_protocols/ms-adts/a29e5c28-9fb9-4c49-8e43-4b9b8e733a05).


## Example

The following code example switches the language of the connection to Spanish.

```vb
Dim myConnection As OLEDBConnection 
Set myConnection = ThisWorkbook.Connections(1) 
 
With myConnection 
 .RetrieveInOfficeUILang = False 
 .LocaleID = 3082 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]