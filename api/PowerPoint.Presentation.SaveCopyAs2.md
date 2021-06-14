---
title: Presentation.SaveCopyAs2 method (PowerPoint)
description: Saves a copy of the specified presentation to a file without modifying the original.
keywords: vbapp10.chm583135
f1_keywords:
- vbapp10.chm583135
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SaveCopyAs2
ms.author: lindalu
ms.date: 09/18/2020
localization_priority: Normal
---


# Presentation.SaveCopyAs2 method (PowerPoint)

Saves a copy of the specified presentation to a file without modifying the original.

## Syntax

_expression_. `SaveCopyAs2`( `_FileName_`, `_FileFormat_`, `_EmbedTrueTypeFonts_`, `_ReadOnlyRecommended_` )

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the name to save the file under. If you don't include a full path, PowerPoint saves the file in the current folder.|
| _FileFormat_|Optional|**PpSaveAsFileType**|The file format.|
| _EmbedTrueTypeFonts_|Optional|**MsoTriState**|Specifies whether TrueType fonts are embedded.|
| _ReadOnlyRecommended_|Optional|**MsoTriState**|Specifies whether the file should be marked as ReadOnlyRecommended.|

## Remarks

The  _FileFormat_ parameter value can be one of these **PpSaveAsFileType** constants. The default is **ppSaveAsDefault**.

||
|:-----|
|**ppSaveAsHTMLv3**|
|**ppSaveAsAddIn**|
|**ppSaveAsBMP**|
|**ppSaveAsDefault**|
|**ppSaveAsGIF**|
|**ppSaveAsHTML**|
|**ppSaveAsHTMLDual**|
|**ppSaveAsJPG**|
|**ppSaveAsMetaFile**|
|**ppSaveAsPNG**|
|**ppSaveAsPowerPoint3**|
|**ppSaveAsPowerPoint4**|
|**ppSaveAsPowerPoint4FarEast**|
|**ppSaveAsPowerPoint7**|
|**ppSaveAsPresentation**|
|**ppSaveAsRTF**|
|**ppSaveAsShow**|
|**ppSaveAsTemplate**|
|**ppSaveAsTIF**|
|**ppSaveAsWebArchive**|

The _EmbedTrueTypeFonts_ parameter value can be one of these **MsoTriState** constants.

|Constant|Description|
|:-----|:-----|
|**msoFalse**|TrueType fonts are not embedded.|
|**msoTriStateMixed**|Embedded fonts are a mixture of TrueType and non-TrueType. The default. |
|**msoTrue**|TrueType fonts are embedded.|

The _ReadOnlyRecommended_ parameter value can be one of these **MsoTriState** constants.

|Constant|Description|
|:-----|:-----|
|**msoFalse**|The new file is not marked as ReadOnlyRecommended.|
|**msoTriStateMixed**|The new file will have the same ReadOnlyRecommended state as the original. The default. |
|**msoTrue**|The new file will be marked as ReadOnlyRecommended. |

## Example

This example saves a copy of the active presentation under the name "New File.pptx" while removing the ReadOnlyRecommended flag.  

```vb
With Application.ActivePresentation

    .SaveCopyAs2 "New File", ppSaveAsDefault, msoTriStateMixed, msoFalse

End With
```

## See also

[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
