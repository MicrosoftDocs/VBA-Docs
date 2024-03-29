---
title: AddIns2.Add method (Excel)
keywords: vbaxl10.chm867073
f1_keywords:
- vbaxl10.chm867073
api_name:
- Excel.AddIns2.Add
ms.assetid: c313e123-9917-f002-bded-cff50085002b
ms.date: 04/03/2019
ms.localizationpriority: medium
---


# AddIns2.Add method (Excel)

Adds a new add-in to the list of add-ins.


## Syntax

_expression_.**Add** (_FileName_, _CopyFile_)

_expression_ A variable that returns an **[AddIns2](Excel.AddIns2.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file that contains the add-in to add to the list in the **Add-Ins** dialog box.|
| _CopyFile_|Optional| **Variant**| If the add-in file is on a removable medium, specifies whether to copy the add-in to the local hard disk. Specify **True** to copy the add-in to your hard disk. Specify **False** to keep the add-in on the removable medium.<br/><br/> If this argument is omitted, Microsoft Excel displays a dialog box and asks the user to choose whether to copy the add-in file. This parameter is ignored if the add-in file is already on the hard disk.|

## Return value

AddIn



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]