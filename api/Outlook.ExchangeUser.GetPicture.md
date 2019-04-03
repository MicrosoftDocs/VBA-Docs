---
title: ExchangeUser.GetPicture method (Outlook)
keywords: vbaol11.chm3485
f1_keywords:
- vbaol11.chm3485
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetPicture
ms.assetid: 4298db85-0576-4982-9592-6eae666d966a
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.GetPicture method (Outlook)

Obtains an  **[IPictureDisp](https://docs.microsoft.com/windows/desktop/api/ocidl/nn-ocidl-ipicturedisp)** object that represents the picture of the Microsoft Exchange user that is displayed in Microsoft Outlook.


## Syntax

_expression_.**GetPicture**

_expression_ A variable that represents an **[ExchangeUser](Outlook.ExchangeUser.md)** object.


## Return value

An  **IPictureDisp** object that represents the picture of the Exchange user that is displayed in Outlook.


## Remarks

The picture of the Exchange user is stored in Active Directory and displayed in various places in Outlook, including the dialog box for  **Outlook Properties** and Contact Card.

If the picture does not exist for the user,  **GetPicture** returns **Null** (**Nothing** for Visual Basic).

You can only call  **GetPicture** from code that runs in-process as Outlook. An **StdPicture** object cannot be marshaled across process boundaries. If you attempt to call **GetPicture** from out-of-process code, an exception occurs. 





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]