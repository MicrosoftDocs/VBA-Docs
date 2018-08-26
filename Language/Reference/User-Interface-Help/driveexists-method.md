---
title: DriveExists Method
keywords: vblr6.chm2182038
f1_keywords:
- vblr6.chm2182038
ms.prod: office
api_name:
- Office.DriveExists
ms.assetid: ddba70e5-8b60-4ce6-631f-fb10f81a6d93
ms.date: 06/08/2017
---


# DriveExists Method



 **Description**
Returns  **True** if the specified drive exists; **False** if it does not.
<<<<<<< HEAD
 **Syntax**
 _object_. **DriveExists(**_drivespec_**)**
=======

## Syntax

_object_. **DriveExists(**_drivespec_**)**
>>>>>>> master
The  **DriveExists** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. Always the name of a  **FileSystemObject**.|
| _drivespec_|Required. A drive letter or a path specification for the root of the drive.|

<<<<<<< HEAD
 **Remarks**
=======
## Remarks

>>>>>>> master
For drives with removable media, the  **DriveExists** method returns **True** even if there are no media present. Use the **IsReady** property of the **Drive** object to determine if a drive is ready.

