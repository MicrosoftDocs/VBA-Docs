---
title: ApplicationSettings.DrawingBackgroundColorGradient property (Visio)
keywords: vis_sdr.chm16251805
f1_keywords:
- vis_sdr.chm16251805
ms.prod: visio
api_name:
- Visio.ApplicationSettings.DrawingBackgroundColorGradient
ms.assetid: 3bd4693b-4312-3b99-5f48-a4d7909cf41c
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.DrawingBackgroundColorGradient property (Visio)

Determines the background gradient color of the Microsoft Visio drawing window for the current session. Read/write. 


## Syntax

_expression_.**DrawingBackgroundColorGradient**

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

Setting the  **BackgroundColorGradient** property of the active window to a value other than the default (-1) overrides the **DrawingBackgroundColorGradient** setting for that window. To be able to reset the background gradient color of the same active window by setting the **DrawingBackgroundColorGradient** property, you must reset **BackgroundColorGradient** to its default value, -1. If multiple windows are open, setting **BackgroundColorGradient** for one window has no effect on the setting for other open windows.




> [!NOTE] 
> You can specify two colors for the drawing background. If users' screen resolution is adequate, one of the colors will grade into the other from the top to the bottom of the screen. To be able to use this feature, users must set their monitors to display 32-bit color. The ability to set drawing background color programmatically for users running in high-contrast mode is restricted.


## Example

The following VBA macro shows how to use the  **DrawingBackgroundColorGradient** property to get and set the application background gradient color. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object, and it demonstrates the relationship between the **DrawingBackgroundColorGradient** property and the **Window.BackgroundColorGradient** property. This example assumes there is a drawing window open in Visio and that initially all background gradient color properties are set to their default values.


```vb
Public Sub DrawingBackgroundColorGradient_Example() 
 
 Dim vsoApplicationSettings As Visio.ApplicationSettings 
 Set vsoApplicationSettings = Visio.Application.Settings 
 
 'Get the current application background gradient color. 
 Debug.Print vsoApplicationSettings.DrawingBackgroundColorGradient 
 
 'Get the active window background color gradient. 
 Debug.Print ActiveWindow.BackgroundColorGradient 
 
 'Change the application background gradient color. 
 'This also changes the active window color as 
 'well as the setting in the Color Settings dialog box. 
 vsoApplicationSettings.DrawingBackgroundColor = vbRed 
 
 'Change the active window background gradient color. 
 ActiveWindow.BackgroundColorGradient = vbMagenta 
 
 'Change the application background gradient color again. 
 'This time, there is no change in the current 
 'window color, but the dialog box setting changes. 
 vsoApplicationSettings.DrawingBackgroundColorGradient = vbYellow 
 
 'Reset Window.BackgroundColorGradient to its default value. 
 ActiveWindow.BackgroundColorGradient = -1 
 
 'Change the application background gradient color again. 
 'Now both the active window color 
 'and the dialog box setting change. 
 vsoApplicationSettings.DrawingBackgroundColorGradient = vbBlue 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]