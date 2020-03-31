---
title: AllowSpecialKeys property
ROBOTS: INDEX
keywords: vbaac10.chm4266
f1_keywords:
- vbaac10.chm4266
ms.prod: access
api_name:
- Access.AllowSpecialKeys
ms.assetid: 5628e6b6-f253-a435-5bce-58b6727382cc
ms.date: 06/08/2019
localization_priority: Normal
---


# AllowSpecialKeys property

**Applies to:** Access 2013 | Access 2016

You can use the **AllowSpecialKeys** property to specify whether or not special key sequences (ALT+F1 (F11), CTRL+F11, CTRL+BREAK, and CTRL+G) are disabled or enabled. For example, you can use the **AllowSpecialKeys** property to prevent a user from displaying the Database window by pressing F11, entering break mode within a Visual Basic module by pressing CTRL+BREAK, or displaying the Immediate window by pressing CTRL+G.


## Setting

The **AllowSpecialKeys** property uses the following settings.

|Setting|Description|
|:-----|:-----|
|**True** (-1)|Enable the special key sequences.|
|**False** (0)|Disable the special key sequences.|

The easiest way to set this property is by using the **Use Access Special Keys** option in the **Current Database** section of the **Access Options** dialog box. 

To view the **Access Options** dialog box, click the **Microsoft Office button**
![File menu button](../../../images/O12FileMenuButton_ZA10077102.gif), and then click **Access Options**. In a Microsoft Access database, you can also set this property by using a macro or Visual Basic.

To set the **AllowSpecialKeys** property by using a macro or Visual Basic, you must first either set the property in the **Access Options** dialog box once or create the property in the following ways:

- In a Microsoft Access database, you can add it by using the **[CreateProperty](https://msdn.microsoft.com/library/f2039be9-5fd8-f673-dfbf-0a71540cdc98%28Office.15%29.aspx)** method and append it to the **Properties** collection of the **Database** object.
    
- In a Microsoft Access project (.adp), you can add it to the **[AccessObjectProperties](https://msdn.microsoft.com/library/2df86891-6038-d147-2a32-f1c77b841067%28Office.15%29.aspx)** collection of the **[CurrentProject](https://msdn.microsoft.com/library/e6baae73-1eeb-b48f-d35e-b3e921378561%28Office.15%29.aspx)** object by using the **[Add](https://msdn.microsoft.com/library/8f86d5f8-b9af-87d3-fae4-e1a24d7225b6%28Office.15%29.aspx)** method.
    

## Remarks

You should make sure the **AllowSpecialKeys** property is set to **True** when debugging an application.

The **AllowSpecialKeys** property affects the following key sequences.

|**Key sequences**|**Effect**|
|:-----|:-----|
|ALT+F1 (F11)|Display the Navigation Pane.|
|CTRL+G|Display the Immediate window.|
|CTRL+F11|Toggle between the custom menu bar and the built-in menu bar.|
|CTRL+BREAK|Enter break mode and display the current module in the Code window.|

This property's setting doesn't take effect until the next time the application database opens.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]