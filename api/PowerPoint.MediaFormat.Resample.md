---
title: MediaFormat.Resample method (PowerPoint)
keywords: vbapp10.chm724013
f1_keywords:
- vbapp10.chm724013
api_name:
- PowerPoint.MediaFormat.Resample
ms.assetid: d1bb8b41-4640-c57c-83bc-3263376b425e
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# MediaFormat.Resample method (PowerPoint)

Adds the current media object to the queue and begins resampling, based on the specified parameters.


## Syntax

_expression_. `Resample`( `_Trim_`, `_SampleHeight_`, `_SampleWidth_`, `_VideoFrameRate_`, `_AudioSamplingRate_`, `_VideoBitRate_` )

 _expression_ An expression that returns a [MediaFormat](PowerPoint.MediaFormat.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Trim_|Optional|**Boolean**|Whether to trim the sample.|
| _SampleHeight_|Optional|**Integer**|The sample resolution height.|
| _SampleWidth_|Optional|**Integer**|The sample resolution width.|
| _VideoFrameRate_|Optional|**Long**|The video frame rate, in frames per second.|
| _AudioSamplingRate_|Optional|**Long**|The audio sampling rate, in bits per second.|
| _VideoBitRate_|Optional|**Long**|The video bit rate, in bits per second.|

## Return value

Nothing


## Remarks

 **Resample** ignores the values of parameters that are not applicable to the media.

If the media content is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## See also


[MediaFormat Object](PowerPoint.MediaFormat.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]