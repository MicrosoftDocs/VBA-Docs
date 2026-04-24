---
title: AcSysCmdAction enumeration (Access)
keywords: vbaac10.chm10027
f1_keywords:
- vbaac10.chm10027
api_name:
- Access.AcSysCmdAction
ms.assetid: a2879d50-9845-40b0-9e51-a022340c664b
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# AcSysCmdAction enumeration (Access)

Used with the **SysCmd** method to specify an action to take.

|Name|Value|Description|
|:-----|:-----|:-----|
|**acSysCmdAccessDir**|9|Returns the name of the directory where Msaccess.exe is located.|
|**acSysCmdAccessVer**|7|Returns the version number of Microsoft Access.|
|**acSysCmdClearHelpTopic**|11|Resets default help topic.|
|**acSysCmdClearStatus**|5|Provides information on the state of a database object.|
|**acSysCmdCompile**|603|Compiles the Visual Basic code modules in the current database. Equivalent to the **Debug > Compile** menu command.|
|**acSysCmdGetBitness**|724|Returns `"32-bit"` or `"64-bit"` as a string matching the bitness of the running binary. Version 2604 and later.|
|**acSysCmdGetBuildNumber**|725|Returns the major build number (for example, `19916`) as a **Long**. Version 2604 and later.|
|**acSysCmdGetChannelName**|723|Returns the update channel name (for example, `"Current Channel"`, `"Monthly Enterprise Channel"`, or `"LTSC 2024"`). Version 2604 and later.|
|**acSysCmdGetFullBuildNumber**|722|Returns the full four-part build string (for example, `"16.0.19916.30000"`). Version 2604 and later.|
|**acSysCmdGetFullVersion**|720|Returns a display string containing version, build, channel, and bitness (for example, `"Microsoft Access (Version 2601) Build 16.0.19628.20000 Current Channel 64-bit"`). Version 2604 and later.|
|**acSysCmdGetMsoBuildNumber**|715|Returns the build number of the shared MSO component as a **Long**. This is the same value returned by **Application.Build**, and it may differ from the Access application build. Use **acSysCmdGetBuildNumber** (725) in new code to get the Access build number.|
|**acSysCmdGetObjectState**|10|Returns the state of the specified database object. You must specify argument1 and argument2 when you use this action value.|
|**acSysCmdGetVersion**|721|Returns the short YYMM marketing version (for example, `"2601"`). Version 2604 and later.|
|**acSysCmdGetWorkgroupFile**|13|Returns the path to the workgroup file (System.mdw).|
|**acSysCmdIniFile**|8|Returns the name of the .ini file associated with Microsoft Access.|
|**acSysCmdInitMeter**|1|Initializes the progress meter. You must specify the argument1 and argument2 arguments when you use this action.|
|**acSysCmdProfile**|12|Returns the **\/profile** setting specified by the user when starting Microsoft Access from the command line.|
|**acSysCmdRemoveMeter**|3|Removes the progress meter.|
|**acSysCmdRuntime**|6|Returns **True** (1) if a run-time version of Microsoft Access is running.|
|**acSysCmdSetStatus**|4|Sets the status bar text to the text argument.|
|**acSysCmdUpdateMeter**|2|Updates the progress meter with the specified value. You must specify the text argument when you use this action.|


## Version, build, and channel information

The following **AcSysCmdAction** constants were added in **Version 2604** to simplify retrieving Access version, build, channel, and bitness information from VBA:

- **acSysCmdGetFullVersion** (720) — composed display string
- **acSysCmdGetVersion** (721) — YYMM marketing version
- **acSysCmdGetFullBuildNumber** (722) — four-part build string
- **acSysCmdGetChannelName** (723) — update channel name
- **acSysCmdGetBitness** (724) — bitness of the running binary
- **acSysCmdGetBuildNumber** (725) — major build number as a **Long**

The display string returned by **acSysCmdGetFullVersion** is intended for display in logs, dialogs, or bug reports. Do not parse it as a structured format; use the individual actions (**acSysCmdGetVersion**, **acSysCmdGetFullBuildNumber**, **acSysCmdGetChannelName**, **acSysCmdGetBitness**) for programmatic access to the components.

### Availability

These action codes are available starting in **Version 2604** of Microsoft 365 Apps. They are **not available** on LTSC 2021 or LTSC 2024, which shipped before these actions existed.

### Example

```vb
Debug.Print SysCmd(acSysCmdGetFullVersion)
' "Microsoft Access (Version 2601) Build 16.0.19628.20000 Current Channel 64-bit"

Debug.Print SysCmd(acSysCmdGetChannelName)
' "Current Channel"

Dim build As Long
build = SysCmd(acSysCmdGetBuildNumber)
' 19916
```

## Previously undocumented action codes

The following **AcSysCmdAction** constants refer to action codes that have existed in Access for some time but were not previously documented. They are now formally named as of **Version 2604**:

- **acSysCmdCompile** (603) — compile the Visual Basic code modules in the current database.
- **acSysCmdGetMsoBuildNumber** (715) — build number of the shared MSO component as a **Long** (same value as **Application.Build**; may differ from the Access build). Use **acSysCmdGetBuildNumber** (725) in new code to get the Access build.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
