---
title: Intrinsic constants as bit masks
ROBOTS: INDEX
keywords: vbaac10.chm13876
f1_keywords:
- vbaac10.chm13876
ms.prod: access
ms.assetid: 2f339b2c-d078-aedd-0ebd-8d04877cbf9a
ms.date: 06/08/2019
localization_priority: Normal
---


# Intrinsic constants as bit masks

**Applies to:** Access 2013 | Access 2016

To test for the Button or Shift arguments, use a bit mask.

The Button argument is a bit field with bits corresponding to the left mouse button (bit 0), right mouse button (bit 1), and middle mouse button (bit 2). These bits correspond to the values 1, 2, and 4, respectively. Only one of the bits is set, indicating which button triggered the event.

The intrinsic constants that Microsoft Access provides for the Button argument have the following values.

|**Constant**|**Value**|
|:-----|:-----:|
|**acLeftButton**|1|
|**acRightButton**|2|
|**acMiddleButton**|4|

The Shift argument is a bit field, with the least significant bits corresponding to the SHIFT key (bit 0), the CTRL key (bit 1), and the ALT key (bit 2 ). These bits correspond to the values 1, 2, and 4, respectively. The Shift argument indicates the state of these keys. Some, all, or none of the bits can be set, indicating that some, all, or none of the keys is pressed. For example, if both CTRL and ALT were pressed, the value of the Shift argument would be 6.

The intrinsic constants that Microsoft Access provides for the Shift argument have the following values.

|**Constant**|**Value**|
|:-----|:-----:|
|**acShiftMask**|1|
|**acCtrlMask**|2|
|**acAltMask**|4|

You can use the constants to test for any combination of buttons and keys without having to figure out the unique bit field value for each combination. A bit is set if the button or key is pressed.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]