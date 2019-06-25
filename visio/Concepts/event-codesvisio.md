---
title: Event codes (Visio)
keywords: vis_sdr.chm81901708
f1_keywords:
- vis_sdr.chm81901708
ms.prod: visio
ms.assetid: de8f5c7a-421d-ebcf-22b6-4310a202ef64
ms.date: 06/24/2019
localization_priority: Normal
---


# Event codes (Visio)

When you are working with the **[Add](../../api/Visio.EventList.Add.md)** or **[AddAdvise](../../api/Visio.EventList.AddAdvise.md)** method, use the following table to find the event code for the event that you want to create. This table lists each Visio event and its corresponding event code and numeric code.

> [!NOTE] 
> If you are using Visual Basic or Visual Basic for Applications (VBA), you don't need to create your own events. See the event topic in this reference that corresponds to the event that you want to use.


## Table of events and corresponding event and numeric codes

|Event|Event code|Numeric code|
|:----|:---------|:-----------|
| **[AfterModal](../../api/Visio.Application.AfterModal.md)**|visEvtApp+visEvtAfterModal|&H1040 (4160)|
| **[AfterResume](../../api/Visio.Application.AfterResume.md)**|visEvtCodeAfterResume|&H00D1 (209)|
| **[AfterResumeEvents](../../api/Visio.Application.AfterResumeEvents.md)**|visEvtCodeAfterResumeEvents|&H00D5 (213)|
| **[AfterRemoveHiddenInformation](../../api/Visio.Application.AfterRemoveHiddenInformation.md)**|visEvtRemoveHiddenInformation|&H000A (11)|
| **[AppActivated](../../api/Visio.Application.AppActivated.md)**|visEvtApp+visEvtAppActivate|&H1001 (4097)|
| **[AppDeactivated](../../api/Visio.Application.AppDeactivated.md)**|visEvtApp+visEvtAppDeactivate|&H1002 (4098)|
| **[AppObjActivated](../../api/Visio.Application.AppObjActivated.md)**|visEvtApp+visEvtObjActivate|&H1004 (4100)|
| **[AppObjDeactivated](../../api/Visio.Application.AppObjDeactivated.md)**|visEvtApp+visEvtObjDeactivate|&H1008 (4104)|
| **[BeforeDataRecordsetDelete](../../api/Visio.Application.BeforeDataRecordsetDelete.md)**|visEvtDel+visEvtDataRecordset|&H4020 (16416)|
| **[BeforeDocumentClose](../../api/Visio.Application.BeforeDocumentClose.md)**|visEvtDel+visEvtDoc|&H4002 (16386)|
| **[BeforeDocumentSave](../../api/Visio.Application.BeforeDocumentSave.md)**|visEvtCodeBefDocSave|&H0007 (7)|
| **[BeforeDocumentSaveAs](../../api/Visio.Application.BeforeDocumentSaveAs.md)**|visEvtCodeBefDocSaveAs|&H0008 (8)|
| **[BeforeMasterDelete](../../api/Visio.Application.BeforeMasterDelete.md)**|visEvtDel+visEvtMaster|&H4008 (16392)|
| **[BeforeModal](../../api/Visio.Application.BeforeModal.md)**|visEvtApp+visEvtBeforeModal|&H1020 (4128)|
| **[BeforePageDelete](../../api/Visio.Application.BeforePageDelete.md)**|visEvtDel+visEvtPage|&H4010 (16400)|
| **[BeforeQuit](../../api/Visio.Application.BeforeQuit.md)**|visEvtApp+visEvtBeforeQuit|&H1010 (4112)|
| **[BeforeSelectionDelete](../../api/Visio.Application.BeforeSelectionDelete.md)**|visEvtCodeBefSelDel|&H0385 (901)|
| **[BeforeShapeDelete](../../api/Visio.Application.BeforeShapeDelete.md)**|visEvtDel+visEvtShape|&H4040 (16448)|
| **[BeforeShapeTextEdit](../../api/Visio.Application.BeforeShapeTextEdit.md)**|visEvtCodeShapeBeforeTextEdit|&H0323 (803)|
| **[BeforeStyleDelete](../../api/Visio.Application.BeforeStyleDelete.md)**|visEvtDel+visEvtStyle|&H4004 (16388)|
| **[BeforeSuspend](../../api/Visio.Application.BeforeSuspend.md)**|visEvtCodeBeforeSuspend|&H00D0 (208)|
| **[BeforeSuspendEvents](../../api/Visio.Application.BeforeSuspendEvents.md)**|visEvtCodeBeforeSuspendEvents|&H00D4 (212)|
| **[BeforeWindowClosed](../../api/Visio.Application.BeforeWindowClosed.md)**|visEvtDel+visEvtWindow|&H4001 (16385)|
| **[BeforeWindowPageTurn](../../api/Visio.Application.BeforeWindowPageTurn.md)**|visEvtCodeBefWinPageTurn|&H02BF (703)|
| **[BeforeWindowSelDelete](../../api/Visio.Application.BeforeWindowSelDelete.md)**|visEvtCodeBefWinSelDel|&H02BE (702)|
| **[CalloutRelationshipAdded](../../api/Visio.Application.CalloutRelationshipAdded.md)**|visEvtCodeCalloutRelationshipAdded|&H01F8 (504)|
| **[CalloutRelationshipDeleted](../../api/Visio.Application.CalloutRelationshipDeleted.md)**|visEvtCodeCalloutRelationshipDeleted|&H01F9 (505)|
| **[CellChanged](../../api/Visio.Application.CellChanged.md)**|visEvtMod+visEvtCell|&H2800 (10240)|
| **[ConnectionsAdded](../../api/Visio.Application.ConnectionsAdded.md)**|visEvtAdd+visEvtConnect|&H8100 (33024)|
| **[ConnectionsDeleted](../../api/Visio.Application.ConnectionsDeleted.md)**|visEvtDel+visEvtConnect|&H4100 (16640)|
| **[ContainerRelationshipAdded](../../api/Visio.Application.ContainerRelationshipAdded.md)**|visEvtCodeContainerRelationshipAdded|&H01F6 (502)|
| **[ContainerRelationshipDeleted](../../api/Visio.Application.ContainerRelationshipDeleted.md)**|visEvtCodeContainerRelationshipDeleted|&H01F7 (503)|
| **[ConvertToGroupCanceled](../../api/Visio.Application.ConvertToGroupCanceled.md)**|visEvtCodeCancelConvertToGroup|&H038C (908)|
| **[DataRecordsetAdded](../../api/Visio.Application.DataRecordsetAdded.md)**|visEvtAdd+visEvtDataRecordset|&H8020 (32800)|
| **[DataRecordsetChanged](../../api/Visio.Application.DataRecordsetChanged.md)**|visEvtMod+VisEvtDataRecordset|&H2020 (8224)|
| **[DesignModeEntered](../../api/Visio.Application.DesignModeEntered.md)**|visEvtCodeDocDesign|&H0006 (6)|
| **DocumentAdded**|visEvtAdd+visEvtDoc|&H8002 (32770)|
| **[DocumentChanged](../../api/Visio.Application.DocumentChanged.md)**|visEvtMod+visEvtDoc|&H2002 (8194)|
| **[DocumentCloseCanceled](../../api/Visio.Application.DocumentCloseCanceled.md)**|visEvtCodeCancelDocClose|&H0010 (10)|
| **[DocumentCreated](../../api/Visio.Application.DocumentCreated.md)**|visEvtCodeDocCreate|&H0001 (1)|
| **[DocumentOpened](../../api/Visio.Application.DocumentOpened.md)**|visEvtCodeDocOpen|&H0002 (2)|
| **[DocumentSaved](../../api/Visio.Application.DocumentSaved.md)**|visEvtCodeDocSave|&H0003 (3)|
| **[DocumentSavedAs](../../api/Visio.Application.DocumentSavedAs.md)**|visEvtCodeDocSaveAs|&H0004 (4)|
| **[EnterScope](../../api/Visio.Application.EnterScope.md)**|visEvtCodeEnterScope|&H00CA (202)|
| **[ExitScope](../../api/Visio.Application.ExitScope.md)**|visEvtCodeExitScope|&H00CB (203)|
| **[FormulaChanged](../../api/Visio.Application.FormulaChanged.md)**|visEvtMod+visEvtFormula|&H3000 (12288)|
| **[GroupCanceled](../../api/Visio.Application.GroupCanceled.md)**|visEvtCodeCancelSelGroup|&H038E (910)|
| **[KeyDown](../../api/Visio.Application.KeyDown.md)**|visEvtCodeKeyDown|&H02C8 (712)|
| **[KeyPress](../../api/Visio.Application.KeyPress.md)**|visEvtCodeKeyPress|&H02C9 (713)|
| **[KeyUp](../../api/Visio.Application.KeyUp.md)**|visEvtCodeKeyUp|&H02CA (714)|
| **[MasterAdded](../../api/Visio.Application.MasterAdded.md)**|visEvtAdd+visEvtMaster|&H8008 (32776)|
| **[MarkerEvent](../../api/Visio.Application.MarkerEvent.md)**|visEvtApp+visEvtMarker|&H1100 (4352)|
| **[MasterChanged](../../api/Visio.Application.MasterChanged.md)**|visEvtMod+visEvtMaster|&H2008 (8200)|
| **[MasterDeleteCanceled](../../api/Visio.Application.MasterDeleteCanceled.md)**|visEvtCodeCancelMasterDel|&H0191 (401)|
| **[MouseDown](../../api/Visio.Application.MouseDown.md)**|visEvtCodeMouseDown|&H02C5 (709)|
| **[MouseMove](../../api/Visio.Application.MouseMove.md)**|visEvtCodeMouseMove|&H02C6 (710)|
| **[MouseUp](../../api/Visio.Application.MouseUp.md)**|visEvtCodeMouseUp|&H02C7 (711)|
| **[MustFlushScopeBeginning](../../api/Visio.Application.MustFlushScopeBeginning.md)**|visEvtCodeBefForcedFlush|&H00C8 (200)|
| **[MustFlushScopeEnded](../../api/Visio.Application.MustFlushScopeEnded.md)**|visEvtCodeAfterForcedFlush|&H00C9 (201)|
| **[NoEventsPending](../../api/Visio.Application.NoEventsPending.md)**|visEvtApp+visEvtNonePending|&H1200 (4608)|
| **[OnKeystrokeMessageForAddon](../../api/Visio.Application.OnKeystrokeMessageForAddon.md)**|visEvtCodeWinOnAddonKeyMSG|&H02C4 (708)|
| **[PageAdded](../../api/Visio.Application.PageAdded.md)**|visEvtAdd+visEvtPage|&H8010 (32784)|
| **[PageChanged](../../api/Visio.Application.PageChanged.md)**|visEvtMod+visEvtPage|&H2010 (8208)|
| **[PageDeleteCanceled](../../api/Visio.Application.PageDeleteCanceled.md)**|visEvtCodeCancelPageDel|&H01F5 (501)|
| **[QueryCancelConvertToGroup](../../api/Visio.Application.QueryCancelConvertToGroup.md)**|visEvtCodeQueryCancelConvertToGroup|&H038B (907)|
| **[QueryCancelDocumentClose](../../api/Visio.Application.QueryCancelDocumentClose.md)**|visEvtCodeQueryCancelDocClose|&H0009 (9)|
| **[QueryCancelGroup](../../api/Visio.Application.QueryCancelGroup.md)**|visEvtCodeQueryCancelSelGroup|&H038D (909)|
| **[QueryCancelMasterDelete](../../api/Visio.Application.QueryCancelMasterDelete.md)**|visEvtCodeQueryCancelMasterDel|&H0190 (400)|
| **[QueryCancelPageDelete](../../api/Visio.Application.QueryCancelPageDelete.md)**|visEvtCodeQueryCancelPageDel|&H01F4 (500)|
| **[QueryCancelQuit](../../api/Visio.Application.QueryCancelQuit.md)**|visEvtCodeQueryCancelQuit|&H00CC (204)|
| **[QueryCancelSelectionDelete](../../api/Visio.Application.QueryCancelSelectionDelete.md)**|visEvtCodeQueryCancelSelDel|&H0387 (903)|
| **[QueryCancelStyleDelete](../../api/Visio.Application.QueryCancelStyleDelete.md)**|visEvtCodeQueryCancelStyleDel|&H012C (300)|
| **[QueryCancelSuspend](../../api/Visio.Application.QueryCancelSuspend.md)**|visEvtCodeQueryCancelSuspend|&H00CE (206)|
| **[QueryCancelSuspendEvents](../../api/Visio.Application.QueryCancelSuspendEvents.md)**|visEvtCodeQueryCancelSuspendEvents|&H0528 (210)|
| **[QueryCancelUngroup](../../api/Visio.Application.QueryCancelUngroup.md)**|visEvtCodeQueryCancelUngroup|&H0389 (905)|
| **[QueryCancelWindowClose](../../api/Visio.Application.QueryCancelWindowClose.md)**|visEvtCodeQueryCancelWinClose|&H02C2 (706)|
| **[QuitCanceled](../../api/Visio.Application.QuitCanceled.md)**|visEvtCodeCancelQuit|&H00CD (205)|
| **[RunModeEntered](../../api/Visio.Application.RunModeEntered.md)**|visEvtCodeDocRunning|&H0005 (5)|
| **[RuleSetValidated](../../api/Visio.Application.RuleSetValidated.md)**|visEvtCodeRuleSetValidated|&H000D (13)|
| **[SelectionAdded](../../api/Visio.Application.SelectionAdded.md)**|visEvtCodeSelAdded|&H0386 (902)|
| **[SelectionChanged](../../api/Visio.Application.SelectionChanged.md)**|visEvtCodeWinSelChange|&H02BD (701)|
| **[SelectionDeleteCanceled](../../api/Visio.Application.SelectionDeleteCanceled.md)**|visEvtCodeCancelSelDel|&H0388 (904)|
| **SelectionMovedToSubprocess**|visEvtCodeSelectionMovedToSubprocess|&H000B (12)|
| **[ShapeAdded](../../api/Visio.Application.ShapeAdded.md)**|visEvtAdd+visEvtShape|&H8040 (32832)|
| **[ShapeChanged](../../api/Visio.Application.ShapeChanged.md)**|visEvtMod+visEvtShape|&H2040 (8256)|
| **[ShapeDataGraphicChanged](../../api/Visio.Application.ShapeDataGraphicChanged.md)**|visEvtShapeDataGraphicChanged|&H0327 (807)|
| **[ShapeExitedTextEdit](../../api/Visio.Application.ShapeExitedTextEdit.md)**|visEvtCodeShapeExitTextEdit|&H0324 (804)|
| **[ShapeLinkAdded](../../api/Visio.Application.ShapeLinkAdded.md)**|visEvtShapeLinkAdded|&H0325 (805)|
| **[ShapeLinkDeleted](../../api/Visio.Application.ShapeLinkDeleted.md)**|visEvtShapeLinkDeleted|&H0326 (806)|
| **[ShapeParentChanged](../../api/Visio.Application.ShapeParentChanged.md)**|visEvtCodeShapeParentChange|&H0322 (802)|
| **ShapesDeleted**|visEvtCodeShapeDelete|&H0321 (801)|
| **[StyleAdded](../../api/Visio.Application.StyleAdded.md)**|visEvtAdd+visEvtStyle|&H8004 (32772)|
| **[StyleChanged](../../api/Visio.Application.StyleChanged.md)**|visEvtMod+visEvtStyle|&H2004 (8196)|
| **[StyleDeleteCanceled](../../api/Visio.Application.StyleDeleteCanceled.md)**|visEvtCodeCancelStyleDel|&H012D (301)|
| **[SuspendCanceled](../../api/Visio.Application.SuspendCanceled.md)**|visEvtCodeCancelSuspend|&H00CF (207)|
| **[SuspendEventsCanceled](../../api/Visio.Application.SuspendEventsCanceled.md)**|visEvtCodeCancelSuspendEvents|&H00D3 (211)|
| **[TextChanged](../../api/Visio.Application.TextChanged.md)**|visEvtMod+visEvtText|&H2080 (8320)|
| **[UngroupCanceled](../../api/Visio.Application.UngroupCanceled.md)**|visEvtCodeCancelUngroup|&H038A (906)|
| **[ViewChanged](../../api/Visio.Application.ViewChanged.md)**|visEvtCodeViewChanged|&H02C1 (705)|
| **[VisioIsIdle](../../api/Visio.Application.VisioIsIdle.md)**|visEvtApp+visEvtIdle|&H1400 (5120)|
| **[WindowActivated](../../api/Visio.Application.WindowActivated.md)**|visEvtApp+visEvtWinActivate|&H1080 (4224)|
| **[WindowCloseCanceled](../../api/Visio.Application.WindowCloseCanceled.md)**|visEvtCodeCancelWinClose|&H02C3 (707)|
| **[WindowOpened](../../api/Visio.Application.WindowOpened.md)**|visEvtAdd+visEvtWindow|&H8001 (32769)|
| **[WindowChanged](../../api/Visio.Application.WindowChanged.md)**|visEvtMod+visEvtWindow|&H2001 (8193)|
| **[WindowTurnedToPage](../../api/Visio.Application.WindowTurnedToPage.md)**|visEvtCodeWinPageTurn|&H02C0 (704)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]