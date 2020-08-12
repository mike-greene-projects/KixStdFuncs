'; Version 1.0  - 2020/07/26 - Mike Greene, Glenn Barnas
';            a - 2020/07/29 - Enable
';            b - 2020/07/29 - Code cleanup
';            c - 2020/07/29 - Virtual configurations
';	     d - 2020/07/30 - Reduced var count. Added ListTrim(). Full test w/ DEBUG=0
';            e - 2020/08/11 - Minor refractor
'; 
'; Arguments:
';	--D             Enable DEBUG mode
';
';       --F             FORCE running all tasks (usually one-time only)
';
';       --S             Schedule the task to run and exit. This is the normal execution method.
';
'; Config Settings
';   ITPConfig.INI
';   Section:PSA_REWRITE
Break On
CLS

SRnd(@TICKS)

' Declare variables
' ======================================================================
Global DEBUG								' Debug flag
Global VERSION								' version string
Global MSG_LOG_, ERR_LOG_						' log filenames - used by fMsg()
Global LOGONLY								' Boolean value for fMsg to suppress screen output
Global APPNAME, APPID							' Application ID (for logging)	

Dim Rc									' Return code catcher
Dim CfgFile								' pathspec for configuration file
Dim Err                                                                ' Error flag - 1+ errors occurred during processing
Dim Force                                                              ' Force flag

Global K_WORK
Global K_MACHGRP
Dim P_Hostname
Dim P_ApiUrl
Dim P_Company
Dim P_PubKey
Dim P_PrivKey
Dim P_AppID
Dim P_Url
Dim P_Auth
Dim BoardId, Board

' ======================================================================
Rc = SetOption("Explicit", "On")
Rc = SetOption("WrapAtEOL", "On")
Rc = SetOption("NoVarsInStrings", "On")
Rc = SetOption("NoMacrosInStrings", "On")
Rc = SetOption("WOW64AlternateRegView", "Off")

' Define values
' ======================================================================
VERSION   = "1.0e"
APPID     = "PSATSI"
MSG_LOG_  = K_WORK + "Logs\" + APPID + "_" + @WDAYNO + "-" + Left(UCase(@DAY), 3) + ".log"
CfgFile   = K_WORK + "ITPConfig.ini"
LOGONLY   = 16
Err       = 0  ' init the error flag
Force     = 0  ' don't force re-running the command
Board     = ReadProfileString($CfgFile, "RMM_SETTINGS", "PSA_CW_Board")

' Set DEBUG and FORCE options
DEBUG     = Bool(InStr(GetCommandLine(0), "--D"))
Force     = Bool(InStr(GetCommandLine(0), "--F"))

' Initialize additional vars here...
Dim aResp
Dim aData, Record, TSI

Dim TypeName, TypeId, TypeIds, aTypes[0], Id, Id1
Dim Subtype, Subtypes, SubtypeId, SubtypeIds, aSubtypes[0], Id2
Dim Item, ItemId, aItems[0], Id3

Dim aId, AIds, aAIds
Dim TypeMap, aTypeMap, SubtypeMap, ItemMap, aItemMap, Map
Dim _, tmpStr, I, J

' Virtual Ini (production), Ini File (Dev/Debug)
' ==========================================================
Dim aIni_T2SI[1, 0], T2SI	' [TYPE]Subtype:Item=1
Dim aIni_I2TS[1, 0], I2TS	' [ITEM]Type:Subtype=1
Dim aIni_IT2S[1, 0], IT2S	' [ITEM:TYPE]Subtype=1
Dim aIni_T2S[1, 0], T2S	        ' [TYPE]Subtype=1
Dim aIni_S2T[1, 0], S2T	        ' [SUBTYPE]Type=1
Dim aIni_Assocs[1, 0], Assocs	' [TYPES]Name/Id LVs, [SUBTYPES]Name/Id LVs,Map_TypeIds, [ITEMS]Name/Id LVs,[ITEM_MAP]TypeId=SubtypeIdsMap	

' ======================================================================
' MAIN CODE - DEFINITIONS
' ======================================================================
' Files
T2SI    = "c:\temp\T2SI.ini"       
I2TS    = "c:\temp\I2TS.ini"	      
IT2S    = "c:\temp\IT2S.ini"
T2S     = "c:\temp\T2S.ini"
S2T     = "c:\temp\S2T.ini"
Assocs  = "c:\temp\assoc.ini"

P_ApiUrl   = "https://&HOSTNAME&/v4_6_release/apis/3.0"
P_Hostname = ReadProfileString(CfgFile, "RMM_SETTINGS", "PSA_URL")
P_Company  = ReadProfileString(CfgFile, "RMM_SETTINGS", "PSA_AV1")
P_PubKey   = ReadProfileString(CfgFile, "RMM_SETTINGS", "PSA_AV2")
P_PrivKey  = ReadProfileString(CfgFile, "RMM_SETTINGS", "PSA_AV3")
' Auth
P_AppID    = ReadProfileString(CfgFile, "RMM_SETTINGS", "PSA_AV4")
P_Url      = Replace(P_ApiUrl, "&HOSTNAME&", P_Hostname)
P_Auth     = cwGenAuth(P_Company, P_PubKey, P_PrivKey, P_Hostname)


' ==========================================================
' MAIN CODE - LOGIC START
' ==========================================================

' If $Board in $CfgFile is Int, use it, else its a name -- look it up!
If Board = Val(Board)
  BoardId = Val(Board)
Else
  BoardId = cwGetBoardIdByName(P_Company, P_PubKey, P_PrivKey, Board)
End If

If $DEBUG
  ' Remove temporary physical configuration files
  Del T2SI
  Del I2TS
  Del IT2S
  Del T2S
  Del S2T
  Del Assocs
End If

' Working data declarations
aData   = EnumIni(CfgFile, "PSA REWRITE")
For Each Record in aData
  TSI = Split(ReadProfileString(CfgFile, "PSA REWRITE", Record), ";")[2]
  ' Virtual configuration
  aIni_T2SI = WriteIniArray(aIni_T2SI, Split(TSI, ",")[0], Split(TSI, ",")[1] + ":" + Split(TSI, ",")[2], 1)
  aIni_I2TS = WriteIniArray(aIni_I2TS, Split(TSI, ",")[2], Split(TSI, ",")[0] + ":" + Split(TSI, ",")[1], 1)
  aIni_IT2S = WriteIniArray(aIni_IT2S, Split(TSI, ",")[2] + ":" + Split(TSI, ",")[0], Split(TSI, ",")[1], 1)
  aIni_T2S = WriteIniArray(aIni_T2S,   Split(TSI, ",")[0], Split(TSI, ",")[1], 1)
  aIni_S2T = WriteIniArray(aIni_S2T,   Split(TSI, ",")[1], Split(TSI, ",")[0], 1)
  
  If $DEBUG
    ' Physical configuration
    _ = WriteProfileString($T2SI, Split($TSI, ",")[0], Split(TSI, ",")[1] + ":" + Split(TSI, ",")[2], 1)
    _ = WriteProfileString($I2TS, Split($TSI, ",")[2], Split(TSI, ",")[0] + ":" + Split(TSI, ",")[1], 1)
    _ = WriteProfileString($IT2S, Split($TSI, ",")[2] + ":" + Split($TSI, ",")[0], Split(TSI, ",")[1], 1)
    _ = WriteProfileString($T2S,  Split($TSI, ",")[0], Split($TSI, ",")[1], 1)
    _ = WriteProfileString($S2T,  Split($TSI, ",")[1], Split($TSI, ",")[0], 1)
  End If
Next

' Enumerate Types
aTypes = EnumIniArray(aIni_T2SI, "")
For Each TypeName in aTypes

  $aResp = cwGetTsiTypeByName($P_Auth, $P_AppID, $P_Url, $BoardId, TypeName)
  If @ERROR
  
   Select
    Case @ERROR = 2 
      ' Create type record - No exact match type could be queried
      aResp = cwCreateTsiType($P_Auth, $P_AppID, $P_Url, $BoardId, TypeName)
      If @ERROR = 400
        fMsg("400: For 'create', exact match type often exists, but exact match query failed.", "", 0, 8 | $LOGONLY, 87)
      End If
      $Id1 = ParseJSON($aResp, "root.Id")
    Case @ERROR = 400
      ' Bad request
      fMsg("400: Check input for parameter error.", "", 0, 8 | $LOGONLY, 87)
    Case @ERROR > 299
      ' Error unknown
      fMsg("Unknown condition.", "", 0, 8 | $LOGONLY, 87)
   End Select
   
  Else
    Id1 = ParseJSON(aResp, "root.Result.0.Id")
  End If
    
  ' Virtual Configuration
  aIni_Assocs = WriteIniArray(aIni_Assocs, "TYPES", TypeName, Id1)
  aIni_Assocs = WriteIniArray(aIni_Assocs, "TYPES", Id1, TypeName)  
  
  If DEBUG
    _ = WriteProfileString(Assocs, "TYPES", TypeName, Id1)
    _ = WriteProfileString(Assocs, "TYPES", Id1, TypeName)
    _ = WriteProfileString(Assocs, "TYPES", "TotalTypes", UBound(aTypes) + 1)  ' DEBUG
  End If
Next

' Enumerate Subtypes, then TYPES, to assign TypeMap for each Subtype
aSubtypes = EnumIniArray(aIni_S2T, "")
For Each Subtype in aSubtypes
  TypeMap = ""  
  ' Enumerate associated types
  aTypes = EnumIniArray(aIni_S2T, Subtype)
  ' Type ID lookup
  For Each TypeName in aTypes
    TypeId = ReadIniArray(aIni_Assocs, "TYPES", TypeName)
    TypeMap = TypeId + "," + TypeMap
  Next
  TypeMap = ListTrim(TypeMap, ",")
  
  ' Virtual Configuration
  aIni_Assocs = WriteIniArray(aIni_Assocs, "SUBTYPES", Subtype + "_MAP", TypeMap)
  
  If DEBUG
    _ = WriteProfileString(Assocs, "SUBTYPES", Subtype + "_MAP", TypeMap)
    If $DEBUG > 1
      ' In the event of major errors, review _MSPB-MAP sanity check
      _ = WriteProfileString(Assocs, "SUBTYPES", $Subtype + "_MSPB-MAP", TypeMap)
    End If
  End If
Next

' Enumerate Subtypes
aSubtypes = EnumIniArray(aIni_S2T, "")
For Each Subtype in aSubtypes
  TypeMap = ReadIniArray(aIni_Assocs, "SUBTYPES", Subtype + "_MAP")
  
  aResp = cwGetTsiSubtypeByName(P_Auth, P_AppID, P_Url, BoardId, Subtype)
  If @ERROR
  
    Select
     Case @ERROR = 2
      ' Create subtype record - No exact match type could be queried
      aResp = cwCreateTsiSubtype(P_Auth, P_AppID, P_Url, BoardId, Subtype, TypeMap)
      If @ERROR = 400 
        fMsg("400: For 'create', exact match type often exists, but exact match query failed.", "", 0, 8 | LOGONLY, 87)
      End If
      Id2 = ParseJSON(aResp, "root.Id")
     Case @ERROR = 400
      ' Bad request
      fMsg("400: Check input for parameter error.", "", 0, 8 | LOGONLY, 87)
     Case @ERROR > 299
      ' Error unknown
      fMsg("Unknown condition.", "", 0, 8 | LOGONLY, 87)
    End Select
   
  Else
    ' Subtype exists
    fMsg("Subtype exists: " + Subtype, "", 0, 4 | LOGONLY, 2)
    Id2  = ParseJSON(aResp, "root.Result.0.Id")
    AIds = ParseJSON(aResp, "root.Result.0.typeAssociationIds")
    
    If AIds
      ' Associations found - need to update MAP
      aIDs = Split(AIds, ",")
      For Each AId in aAIds
        If Not InStr(TypeMap, AId)
          TypeMap = TypeMap + "," + AId
        End If
      Next
    End If
    
    ' Update the associations in ConnectWise - If associations were present, Map was updated above
    aResp = cwUpdateTsiSubtypeAssociations(P_Auth, P_AppID, P_Url, BoardId, Subtype, TypeMap)
    If $DEBUG > 1
      fMsg("T/S/MAP: " + TypeName + "/" + Subtype + "/" + TypeMap, "", 0, 4 | LOGONLY, 2)
    End If
  End If 

  ' Virtual Configuration
  aIni_Assocs = WriteIniArray(aIni_Assocs, "SUBTYPES", Subtype, Id2)
  aIni_Assocs = WriteIniArray(aIni_Assocs, "SUBTYPES", Id2, Subtype)
  
  If $DEBUG
    _ = WriteProfileString($Assocs, "SUBTYPES", Subtype, Id2)
    _ = WriteProfileString($Assocs, "SUBTYPES", Id2, Subtype)
    _ = WriteProfileString($Assocs, "SUBTYPES", "TotalSubtypes", UBound($aSubtypes) + 1)  ' DEBUG
  End If
Next

' Enumerate Items
aItems = EnumIniArray(aIni_I2TS, "")
For Each Item in aItems
  aResp = cwGetTsiItemByName(P_Auth, P_AppID, P_Url, BoardId, Item)
  If @ERROR
  
    Select
     Case @ERROR = 2
      ' Create item record - No exact match type could be queried
      aResp = cwCreateTsiItem(P_Auth, P_AppID, P_Url, BoardId, Item)
      If @ERROR = 400
        fMsg("400: For "create", exact match type often exists, but exact match query failed.", "", 0, 8 | LOGONLY, 87)
      End If
      Id3 = ParseJSON(aResp, "root.Id")
     Case @ERROR = 400
      ' Bad request
      fMsg("400: Check input for parameter error.", "", 0, 8 | LOGONLY, 87)
     Case @ERROR > 299
      ' Error unknown
      fMsg("Unknown condition.", "", 0, 8 | LOGONLY, 87)
    EndSelect
   
  Else
    Id3 = ParseJSON(aResp, "root.Result.0.Id")
  End If 
  
  ' Virtual Configuration
  aIni_Assocs = WriteIniArray(aIni_Assocs, "ITEMS", Item, Id3)
  aIni_Assocs = WriteIniArray(aIni_Assocs, "ITEMS", Id3, Item)
  
  If $DEBUG
    _ = WriteProfileString(Assocs, "ITEMS", Item, Id3)
    _ = WriteProfileString(Assocs, "ITEMS", Id3, Item)
    _ = WriteProfileString(Assocs, "ITEMS", "TotalItems", UBound($aItems) + 1)  ' DEBUG
  End If
Next

' Enumerate [Item:Type] 1:1 relationships (sections)
aData = EnumIniArray(aIni_IT2S, "")
For Each Record in aData
  Item = Split(Record, ":")[0]
  TypeName = Split(Record, ":")[1]
  ItemId = ReadIniArray(aIni_Assocs, "ITEMS", Item)
  TypeId = ReadIniArray(aIni_Assocs, "TYPES", Type)
  Subtypes = EnumIniArray(aIni_IT2S, Record)
  SubtypeMap = ''
  tmpStr = '' ' Use to !sanity_check! -- DEBUG
  
  ' Subtype ID lookup
  For Each Subtype in Subtypes
    SubtypeId = ReadIniArray(aIni_Assocs,"SUBTYPES", Subtype)
    SubtypeMap = SubtypeId + "," + SubtypeMap
    tmpStr = Subtype + "," + tmpStr ' DEBUG - show subtype name
  Next
  ' Remove the trailing comma from the SubtypeMap
  SubtypeMap = ListTrim(SubtypeMap, ",")
  
  ' Virtual Configuration - Write the SubtypeMap: [ITEM]TypeID=111,112,113
  aIni_Assocs = WriteIniArray(aIni_Assocs, "ITEM_" + Item + "_MAP", TypeId, SubtypeMap)
  
  If DEBUG
    _ = WriteProfileString(Assocs, "ITEM_" + Item + "_MAP", TypeId, SubtypeMap)
    If DEBUG > 1
      ' In the event of major errors, review !sanity_check! _Name-map
      _ = WriteProfileString(Assocs, "ITEM_" + Item + "_Name-map", TypeName, tmpStr)
    End If
  End If
Next

' Enumerate Items
aItems = EnumIniArray(aIni_I2TS, "")
For Each Item in aItems
  ItemId = ReadIniArray(aIni_Assocs, "ITEMS", Item)
  aResp = cwGetTsiItemAssociations(P_Auth, P_AppID, P_Url, BoardId, ItemId) 
  If @ERROR = 2
    fMsg("No associations in account for item " + $Item, "", 0, 4 | LOGONLY)
  Else
    TypeIds = ParseJSON(aResp, "root.Result.*.id")
    For J = 0 to UBound(TypeIds)
      TypeName = ReadIniArray(aIni_Assocs, "TYPES", TypeIds[J])
      TypeId = ReadIniArray(aIni_Assocs, "TYPES", TypeName)
      SubtypeMap = ReadIniArray(aIni_Assocs, "ITEM_" + Item + _MAP, TypeIds[J])
      If Not SubtypeMap
        fMsg("No MSPB associations in account for item " + $Item, "", 0, 4 | LOGONLY)
      Else
        SubtypeIds = ParseJSON(aResp, "root.Result." + J + ".subTypeAssociationIds")
        If SubtypeIds
          ' Associations found - need to update MAP
          aIDs = Split(SubtypeIds, ",")
          For Each AId in aAIds
            If Not InStr(SubtypeMap, AId)
            SubtypeMap = SubtypeMap + "," + AId
            End If
          Next
        End If
        ' Virtual Configuration
        aIni_Assocs = WriteIniArray(aIni_Assocs, "ITEM_" + Item + "_MAP", TypeId, SubtypeMap)
        If $DEBUG
           = WriteProfileString(Assocs, "ITEM_" + Item + "_MAP", TypeId, SubtypeMap)
        End If
        fMsg("For type '" + TypeName + "' with item '" + Item + "', found STIds: '" + SubtypeIds, "", 0, 4 | LOGONLY, 2)
      End If
    Next
  End If
Next

' Submit associations for MSPB IDs and MSPB+Current Ids
' Assumes the ITEM_ $ItemName _MAP has been properly updated.
' Enumerate [Item:Type] 1:1 relationships (sections)
aData = EnumIniArray(aIni_IT2S, "")
For Each Record in aData
  Item       = Split(Record, ":")[0]
  TypeName   = Split(Record, ":")[1]
  ItemId     = ReadIniArray(aIni_Assocs, "ITEMS", Item)
  TypeId     = ReadIniArray(aIni_Assocs, "TYPES", TypeName)
  SubtypeMap = ReadIniArray(aIni_Assocs, "ITEM_" + Item + "_MAP", TypeId)
  aResp  = cwAssignTsiItemAssociations(P_Auth, P_AppID, P_Url, BoardId, ItemId, TypeId, SubtypeMap)
Next

' Copy files to avoid direct overwrite in case batch delete required
If DEBUG
  copy Assocs  "c:\temp\assocs_del.ini"
  copy I2TS    "c:\temp\I2TS_del.ini"
  copy T2S     "c:\temp\T2S_del.ini"
  copy S2T     "c:\temp\S2T_del.ini"
End If

Quit 0