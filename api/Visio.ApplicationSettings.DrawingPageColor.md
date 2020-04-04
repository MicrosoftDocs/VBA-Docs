---
title: ApplicationSettings.DrawingPageColor property (Visio)
keywords: vis_sdr.chm16251800
f1_keywords:
- vis_sdr.chm16251800
ms.prod: visio
api_name:
- Visio.ApplicationSettings.DrawingPageColor
ms.assetid: 7ae90e3a-d7ed-588e-2416-eecc3e8619e7
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.DrawingPageColor property (Visio)

Determines the page color of the Microsoft Visio drawing window for the current session. Read/write. 


## Syntax

_expression_.**DrawingPageColor**

_expression_ A variable that represents an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

OLE_COLOR


## Remarks

Valid values for an **OLE_COLOR** property within Visio can be one of the following:




- &H00 _bbggrr,_ where _bb_ is the blue value between 0 and 0xFF (255), _gg_ the green value, and _rr_ the red value.
    
- &H800000 _xx_ , where _xx_ is a valid **GetSysColor** index.
    


For details about the  **GetSysColor** function, search for " **GetSysColor** " in the Microsoft Platform SDK on MSDN.

The  **OLE_COLOR** data type is used for properties that return colors. When a property is declared as **OLE_COLOR**, the Properties window displays a color-picker dialog box that allows the user to select the color for the property visually, rather than having to remember the numeric equivalent.

In addition, you can use the following Microsoft Visual Basic for Applications (VBA) color constants for  **OLE_COLOR**.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **vbBlack**|0x0 |Black|
| **vbRed**|0xFF |Red|
| ** vbGreen**|0xFF00 |Green|
| **vbYellow**|0xFFFF|Yellow|
| **vbBlue**|0xFF0000 |Blue|
| ** vbMagenta**|0xFF00FF |Magenta|
| ** vbCyan**|0xFFFF00|Cyan|
| ** vbWhite**|0xFFFFFF|White|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]