---
title: Options.DefaultFilePath property (Word)
keywords: vbawd10.chm162988097
f1_keywords:
- vbawd10.chm162988097
ms.prod: word
api_name:
- Word.Options.DefaultFilePath
ms.assetid: 39c90157-1824-55ee-c7e1-3687f132131f
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.DefaultFilePath property (Word)

Returns or sets default folders for items such as documents, templates, and graphics. Read/write  **String**.


## Syntax

_expression_. `DefaultFilePath`( `_Path_` )

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **WdDefaultFilePath**|The default folder to set.|

## Remarks

 You can use an empty string ("") to remove the setting from the Windows registry. The new setting takes effect immediately.


## Example

This example sets the default folder for Word documents.


```vb
Options.DefaultFilePath(wdDocumentsPath) = "C:\Documents"
```

This example returns the current default path for user templates (corresponds to the default path setting on the  **File Locations** tab in the **Options** dialog box).




```vb
Dim strPath As String 
 
strPath = Options.DefaultFilePath(wdUserTemplatesPath)
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]