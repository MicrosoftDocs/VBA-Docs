---
title: AppointmentItem.CopyTo method (Outlook)
keywords: vbaol11.chm3517
f1_keywords:
- vbaol11.chm3517
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.CopyTo
ms.assetid: 50b8e820-fdb9-1ee9-289b-99be037300c4
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.CopyTo method (Outlook)

Copies the  **[AppointmentItem](Outlook.AppointmentItem.md)** to the folder that is specified by the _DestinationFolder_ parameter and returns an object that represents the item created in the destination folder by the copy operation.


## Syntax

_expression_. `CopyTo`( `_DestinationFolder_` , `_CopyOptions_` )

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DestinationFolder_|Required| **[Folder](Outlook.Folder.md)**|Specifies the folder to which the  **AppointmentItem** object is copied.|
| _CopyOptions_|Required| **[OlAppointmentCopyOptions](Outlook.OlAppointmentCopyOptions.md)**|Specifies the user experience of the copy operation.|

## Return value

Returns an **AppointmentItem** that represents the object created in the destination folder as a result of the copy operation.


## Remarks

If no argument is specified for the  _CopyOptions_ parameter, **CopyTo** assumes that the value is **olCreateAppointment**.

 **CopyTo** returns an error if the destination folder is not an appropriate folder type for an **AppointmentItem** object, or if the user does not have the necessary permissions to create items in the specified destination folder.

Setting the REG_MULTI_SZ value,  `DisableCrossAccountCopy`, in  `HKCU\Software\Microsoft\Office\15.0\Outlook` in the Windows registry has the side effect of disabling this method.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]