---
title: Page.Picture Property (Outlook Forms Script)
api_name:
- Outlook.page.picture
ms.assetid: 447a0372-d621-9b36-3f62-ad764b7e1b92
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Page.Picture Property (Outlook Forms Script)

Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.


## Syntax

_expression_.**Picture**

_expression_ A variable that represents a **Page** object.


## Remarks

You must use the control's property page to assign a bitmap to the **Picture** property. You cannot use the Visual Basic **LoadPicture** function to assign a bitmap to **Picture**.

To remove a picture that is assigned to a control, click the value of the **Picture** property in the property page and then press **DELETE**. Pressing **BACKSPACE** will not remove the picture.

Use the **[PictureSizeMode](Outlook.page.picturesizemode.md)** property to determine how the picture fills the object.

Transparent pictures sometimes have a hazy appearance. If you don't like this appearance, display the picture on a control that supports opaque images. **[Image](Outlook.image.md)** supports opaque images.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]