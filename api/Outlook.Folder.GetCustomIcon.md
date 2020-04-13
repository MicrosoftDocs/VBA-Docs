---
title: Folder.GetCustomIcon method (Outlook)
keywords: vbaol11.chm3316
f1_keywords:
- vbaol11.chm3316
ms.prod: outlook
api_name:
- Outlook.Folder.GetCustomIcon
ms.assetid: 49a3da64-2b2f-76db-0053-88e35141cca0
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.GetCustomIcon method (Outlook)

Returns an **[IPictureDisp](https://docs.microsoft.com/windows/desktop/api/ocidl/nn-ocidl-ipicturedisp)** object that represents the custom icon for the folder.


## Syntax

_expression_.**GetCustomIcon**

_expression_ A variable that represents a **[Folder](Outlook.Folder.md)** object.


## Return value

An **IPictureDisp** object that represents a custom icon for the folder.


## Remarks

The returned  **IPictureDisp** object has its **Type** property equal to **PICTYPE_ICON** or **PICTYPE_BITMAP**.

**GetCustomIcon** returns **Null** (**Nothing** in Visual Basic) if the folder does not have a custom folder icon, or if the folder belongs to one of the following groups of folders:


- Default folders (as listed by the  **[OlDefaultFolders](Outlook.OlDefaultFolders.md)** enumeration)
    
- Special folders (as listed by the  **[OlSpecialFolders](Outlook.OlSpecialFolders.md)** enumeration)
    
- Exchange public folders
    
- Root folder of any Exchange mailbox
    
- Hidden folders
    
You can only call  **GetCustomIcon** from code that runs in-process as Outlook. An **IPictureDisp** object cannot be marshaled across process boundaries. If you attempt to call **GetCustomIcon** from out-of-process code, an exception occurs. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]