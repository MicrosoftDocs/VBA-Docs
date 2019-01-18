---
title: Shell constants
ms.prod: office
ms.assetid: 76b5cc9e-e896-f658-7d23-ca850305a16b
ms.date: 12/11/2018
localization_priority: Normal
---


# Shell constants

The following [constants](../../Glossary/vbe-glossary.md#constant) can be used anywhere in your code in place of the actual values.

<br/>

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbHide**|0|Window is hidden and focus is passed to the hidden window.|
|**vbNormalFocus**|1|Window has focus and is restored to its original size and position.|
|**vbMinimizedFocus**|2|Window is displayed as an icon with focus.|
|**vbMaximizedFocus**|3|Window is maximized with focus.|
|**vbNormalNoFocus**|4|Window is restored to its most recent size and position. The currently active window remains active.|
|**vbMinimizedNoFocus**|6|Window is displayed as an icon. The currently active window remains active.|

On the Macintosh, **vbNormalFocus**, **vbMinimizedFocus**, and **vbMaximizedFocus** all place the application in the foreground; **vbHide**, **vbNoFocus**, and **vbMinimizedFocus** all place the application in the background.


## See also

- [Constants (Visual Basic for Applications)](../constants-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]