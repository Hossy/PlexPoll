' SPDX-License-Identifier: GPL-3.0-or-later
' Copyright (C) John Greg Hossbach
' PlexPoll is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' PlexPoll is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License for more details.
' Full license text: see repository LICENSE file.

' PlexSessionPoll.vb - Poll Plex /status/sessions and update existing HS devices
' Run from recurring event every 10-30 seconds
' HomeSeer scripts are VB.NET (not VBA/VB6):
' 1) Do not use legacy Let/Set assignment keywords.
' 2) Calls with arguments must use parentheses (example: dict.Add("key", 123)).
' 3) Keep helper signatures explicitly typed to avoid compiler ambiguity.

Sub Main(ByVal parm As Object)

    On Error GoTo ErrorHandler

    Dim DEBUG_LOG As Boolean
    DEBUG_LOG = False
    Dim RUN_ID As String
    RUN_ID = Format(Now, "yyyyMMddHHmmss") & "-" & CStr(CLng(Timer() * 1000) Mod 100000)

    Dim iniFile As String
    iniFile = "PlexSessionPoll.ini"
    Dim blPlexIniFile As String
    blPlexIniFile = "hspi_BLPlex.ini"
    EnsureIniFile(iniFile)

    Dim TOKEN As String
    TOKEN = hs.GetINISetting("Plex", "Token", "", iniFile)
    If TOKEN = "" Or UCase(TOKEN) = "REPLACE_ME_WITH_PLEX_TOKEN" Then
        LogMessage(RUN_ID, "ERROR", "Missing [Plex] Token in " & iniFile)
        Exit Sub
    End If

    Dim SERVER As String
    SERVER = ReadIniStringWithFallback("Plex", "Server", iniFile, "Settings", "plexIpAddress", blPlexIniFile, "")
    If SERVER = "" Or UCase(SERVER) = "REPLACE_ME_WITH_PLEX_SERVER" Then
        LogMessage(RUN_ID, "ERROR", "Missing [Plex] Server in " & iniFile & " (hostname or IP)")
        Exit Sub
    End If

    Dim showLastPlayedWhenStopped As Boolean
    Dim showLastRaw As String
    showLastRaw = ReadIniStringWithFallback("Plex", "showLastPlayedWhenStopped", iniFile, "Settings", "showLastPlayedWhenStopped", blPlexIniFile, "false")
    showLastPlayedWhenStopped = IsIniTrue(showLastRaw)

    Dim portRaw As String
    portRaw = ReadIniStringWithFallback("Plex", "Port", iniFile, "Settings", "plexPort", blPlexIniFile, "32400")
    Dim portNumber As Integer
    If Not IsNumeric(portRaw) Then
        LogMessage(RUN_ID, "ERROR", "Port must be numeric. Got '" & portRaw & "'")
        Exit Sub
    End If
    portNumber = CInt(portRaw)
    If portNumber < 1 Or portNumber > 65535 Then
        LogMessage(RUN_ID, "ERROR", "[Plex] Port out of range (1-65535) in " & iniFile)
        Exit Sub
    End If

    Dim USE_HTTPS As Boolean
    USE_HTTPS = IsIniTrue(hs.GetINISetting("Plex", "UseHttps", "False", iniFile))

    Dim defaultTimeoutMs As Integer
    Dim resolveTimeoutMs As Integer
    Dim connectTimeoutMs As Integer
    Dim sendTimeoutMs As Integer
    Dim receiveTimeoutMs As Integer
    defaultTimeoutMs = ReadIniInt("Plex", "TimeoutMs", 10000, iniFile)
    resolveTimeoutMs = ReadIniInt("Plex", "ResolveTimeoutMs", defaultTimeoutMs, iniFile)
    connectTimeoutMs = ReadIniInt("Plex", "ConnectTimeoutMs", defaultTimeoutMs, iniFile)
    sendTimeoutMs = ReadIniInt("Plex", "SendTimeoutMs", defaultTimeoutMs, iniFile)
    receiveTimeoutMs = ReadIniInt("Plex", "ReceiveTimeoutMs", defaultTimeoutMs, iniFile)
    If resolveTimeoutMs < 1000 Then resolveTimeoutMs = 1000
    If connectTimeoutMs < 1000 Then connectTimeoutMs = 1000
    If sendTimeoutMs < 1000 Then sendTimeoutMs = 1000
    If receiveTimeoutMs < 1000 Then receiveTimeoutMs = 1000

    Dim PMS_URL As String
    Dim scheme As String
    scheme = "http"
    If USE_HTTPS Then scheme = "https"
    PMS_URL = scheme & "://" & SERVER & ":" & CStr(portNumber) & "/status/sessions?X-Plex-Token=" & TOKEN

    ' Supported Plex players mapped to HomeSeer parent device refs from [Players] section
    Dim playerMap
    playerMap = LoadPlayerMap(iniFile, RUN_ID)
    If playerMap.Count = 0 Then
        LogMessage(RUN_ID, "ERROR", "No player mappings found in [Players] section of " & iniFile)
        Exit Sub
    End If

    Dim featureMapCache
    featureMapCache = CreateObject("Scripting.Dictionary")
    ValidateMappedFeatures(RUN_ID, playerMap, featureMapCache)

    Dim xmlDoc
    Dim xmlHttp

    On Error GoTo ErrorHandler

    xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.setTimeouts(resolveTimeoutMs, connectTimeoutMs, sendTimeoutMs, receiveTimeoutMs)
    
    xmlHttp.Open("GET", PMS_URL, False)
    xmlHttp.Send
    
    If xmlHttp.Status <> 200 Then
        LogMessage(RUN_ID, "ERROR", "Bad status from Plex: " & xmlHttp.Status)
        GoTo Cleanup
    End If
    
    xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.loadXML(xmlHttp.responseText)
    
    If xmlDoc.parseError.errorCode <> 0 Then
        LogMessage(RUN_ID, "ERROR", "XML parse error: " & xmlDoc.parseError.reason)
        GoTo Cleanup
    End If
    
    Dim sessions
    sessions = xmlDoc.selectNodes("//MediaContainer/Video | //MediaContainer/Track")

    Dim seenPlayers
    seenPlayers = CreateObject("Scripting.Dictionary")
    
    If sessions.length = 0 Then
        DebugLog(DEBUG_LOG, RUN_ID, "PlexPoll", "Got 0 session(s) from Plex")
    End If
    
    DebugLog(DEBUG_LOG, RUN_ID, "PlexPoll", "Got " & sessions.Length & " session(s) from Plex")

    Dim session
    For Each session In sessions
        
        Dim sessionKey As String
        sessionKey = SafeAttr(session, "sessionKey")
        DebugLog(DEBUG_LOG, RUN_ID, "PlexPoll", "Processing sessionKey " & sessionKey)
        
        Dim playerNode
        playerNode = session.selectSingleNode("Player")
        If playerNode Is Nothing Then
            LogMessage(RUN_ID, "WARN", "Failed to get Player node for sessionKey " & sessionKey)
            GoTo NextSession
        End If
        
        Dim playerID As String
        playerID = SafeAttr(playerNode, "machineIdentifier")
        
        Dim state As String
        state = SafeAttr(playerNode, "state")
        If state = "" Then
            state = "Unknown"
        End If
        
        Dim title As String
        title = SafeAttr(session, "title")
        If title = "" Then title = "(untitled)"
        
        Dim duration As String
        duration = SafeAttr(session, "duration")
        
        Dim viewOffset As String
        viewOffset = SafeAttr(session, "viewOffset")
        
        DebugLog(DEBUG_LOG, RUN_ID, "PlexPoll", "Got all data for sessionKey " & sessionKey)

        Dim progress As Double
        progress = 0

        Dim durationMs As Double
        Dim viewOffsetMs As Double
        durationMs = 0
        viewOffsetMs = 0

        If Double.TryParse(duration, durationMs) AndAlso Double.TryParse(viewOffset, viewOffsetMs) AndAlso durationMs > 0 Then
            progress = (viewOffsetMs / durationMs) * 100
        End If
        
        If playerMap.Exists(playerID) Then
            Dim parentRef As Integer
            parentRef = CInt(playerMap(playerID))
            If Not seenPlayers.Exists(playerID) Then seenPlayers.Add(playerID, True)

            Dim featureMap
            featureMap = GetFeatureMapForParent(featureMapCache, parentRef)

            Dim rating As String
            rating = SafeAttr(session, "audienceRating")
            If rating = "" Then
                Dim ratingNode
                ratingNode = session.selectSingleNode("Rating[@type='audience']")
                If Not ratingNode Is Nothing Then
                    rating = SafeAttr(ratingNode, "value")
                End If
            End If

            Dim contentRating As String
            contentRating = SafeAttr(session, "contentRating")

            Dim mediaType As String
            mediaType = SafeAttr(session, "type")

            Dim mediaFilePath As String
            mediaFilePath = ""
            Dim partNode
            partNode = session.selectSingleNode("Media/Part")
            If Not partNode Is Nothing Then
                mediaFilePath = SafeAttr(partNode, "file")
            End If

            Dim mediaFile As String
            mediaFile = FileNameFromPath(mediaFilePath)

            Dim artworkThumb As String
            artworkThumb = SafeAttr(session, "thumb")

            Dim userName As String
            userName = ""
            Dim userNode
            userNode = session.selectSingleNode("User")
            If Not userNode Is Nothing Then
                userName = SafeAttr(userNode, "title")
            End If

            Dim artist As String
            Dim album As String
            artist = ""
            album = ""
            If LCase(mediaType) = "track" Then
                artist = SafeAttr(session, "grandparentTitle")
                album = SafeAttr(session, "parentTitle")
            End If

            Dim durationText As String
            Dim progressText As String
            Dim remainingText As String
            durationText = ToBlPlexDurationText(duration)
            progressText = ToBlPlexProgressText(viewOffset)
            remainingText = ToBlPlexTimeRemainingText(duration, viewOffset)

            SetFeatureString(featureMap, "PlayerIdentifier", playerID)
            SetFeatureString(featureMap, "Rating", rating)
            SetFeatureString(featureMap, "ContentRating", contentRating)
            SetFeatureString(featureMap, "Duration", durationText)
            SetFeatureString(featureMap, "Progress", progressText)
            SetFeatureString(featureMap, "TimeRemaining", remainingText)
            SetFeatureString(featureMap, "MediaType", mediaType)
            SetFeatureString(featureMap, "MediaTitle", title)
            SetFeatureString(featureMap, "MediaFile", mediaFile)
            SetFeatureString(featureMap, "MediaFilePath", mediaFilePath)
            SetFeatureString(featureMap, "ArtworkThumb", artworkThumb)
            SetFeatureString(featureMap, "Artist", artist)
            SetFeatureString(featureMap, "Album", album)
            SetFeatureString(featureMap, "User", userName)

            SetParentStateByValue(parentRef, state)

            If featureMap.Exists("progress") Then
                SetDeviceValueIfChanged(CInt(featureMap("progress")), progress)
            Else
                LogMessage(RUN_ID, "WARN", "Progress feature not found under parent ref " & parentRef)
            End If
        Else
            DebugLog(DEBUG_LOG, RUN_ID, "PlexPoll", "Ignoring session for " & playerID)
        End If
        
NextSession:
    Next

    Dim mappedPlayerID
    For Each mappedPlayerID In playerMap.Keys
        If Not seenPlayers.Exists(CStr(mappedPlayerID)) Then
            ApplyStoppedDefaults(DEBUG_LOG, RUN_ID, CStr(mappedPlayerID), CInt(playerMap(mappedPlayerID)), featureMapCache, showLastPlayedWhenStopped)
        End If
    Next
    
Cleanup:
    xmlDoc = Nothing
    xmlHttp = Nothing
    Exit Sub

ErrorHandler:
    LogMessage(RUN_ID, "ERROR", "Error in polling: " & Err.Description)
    Resume Cleanup

End Sub

Function SafeAttr(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next

    Dim valueObj
    valueObj = node.getAttribute(attrName)

    If IsNothing(valueObj) Or IsDBNull(valueObj) Then
        SafeAttr = ""
    Else
        SafeAttr = CStr(valueObj)
    End If

    On Error GoTo 0
End Function

Sub EnsureIniFile(ByVal iniFile As String)
    On Error Resume Next

    Dim fso
    Dim iniPath As String
    Dim mustInitialize As Boolean
    fso = CreateObject("Scripting.FileSystemObject")
    iniPath = hs.GetAppPath & "\Config\" & iniFile
    mustInitialize = False

    If Not fso.FileExists(iniPath) Then
        mustInitialize = True
    Else
        Dim fi
        fi = fso.GetFile(iniPath)
        If CLng(fi.Size) = 0 Then
            mustInitialize = True
        End If
    End If

    If mustInitialize Then
        Dim tf
        tf = fso.CreateTextFile(iniPath, True, False)
        tf.WriteLine("; PlexSessionPoll configuration")
        tf.WriteLine(";")
        tf.WriteLine("; 1) Sign in to https://app.plex.tv")
        tf.WriteLine("; 2) Open: https://plex.tv/api/resources?includeHttps=1")
        tf.WriteLine("; 3) In browser dev tools, copy your X-Plex-Token from the request URL")
        tf.WriteLine("; 4) Set Server to Plex server hostname or IP (examples: plex.local or 192.168.1.20)")
        tf.WriteLine("; 5) Set Port (default Plex is 32400) and optional UseHttps/TimeoutMs")
        tf.WriteLine("; 6) Add player mappings under [Players]: machineIdentifier=ParentDeviceRef")
        tf.WriteLine("; 7) Paste your token below and save this file")
        tf.WriteLine(";")
        tf.WriteLine("[Plex]")
        tf.WriteLine("Server=REPLACE_ME_WITH_PLEX_SERVER")
        tf.WriteLine("Port=32400")
        tf.WriteLine("UseHttps=False")
        tf.WriteLine("showLastPlayedWhenStopped=false")
        tf.WriteLine("TimeoutMs=10000")
        tf.WriteLine("ResolveTimeoutMs=10000")
        tf.WriteLine("ConnectTimeoutMs=10000")
        tf.WriteLine("SendTimeoutMs=10000")
        tf.WriteLine("ReceiveTimeoutMs=10000")
        tf.WriteLine("Token=REPLACE_ME_WITH_PLEX_TOKEN")
        tf.WriteLine("")
        tf.WriteLine("[Players]")
        tf.WriteLine("; machineIdentifier=ParentDeviceRef")
        tf.WriteLine("; 0f704be1a49fa5bee07af310cf52d9fa=273")
        tf.WriteLine("; e8126407fd12a306-com-plexapp-android=459")
        tf.WriteLine("; 5dc5677ac39af7e8-com-plexapp-android=556")
        tf.Close()

        hs.WriteLog("PlexPoll", "WARN: Initialized INI file at " & iniPath & ". Set [Plex] Token before next run.")
    End If

    On Error GoTo 0
End Sub

Function LoadPlayerMap(ByVal iniFile As String, ByVal runId As String) As Object
    Dim map
    map = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Dim iniPath As String
    iniPath = hs.GetAppPath & "\Config\" & iniFile

    Dim fso
    fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(iniPath) Then
        LoadPlayerMap = map
        On Error GoTo 0
        Exit Function
    End If

    Dim ts
    ts = fso.OpenTextFile(iniPath, 1, False)
    Dim inPlayersSection As Boolean
    Dim lineNo As Integer
    inPlayersSection = False
    lineNo = 0

    Do While Not ts.AtEndOfStream
        Dim line As String
        line = Trim(ts.ReadLine())
        lineNo = lineNo + 1
        If line <> "" Then
            If Left(line, 1) = ";" Or Left(line, 1) = "#" Then
                ' comment
            ElseIf Left(line, 1) = "[" And Right(line, 1) = "]" Then
                inPlayersSection = (LCase(line) = "[players]")
            ElseIf inPlayersSection Then
                Dim eqPos As Integer
                eqPos = InStr(1, line, "=", vbTextCompare)
                If eqPos > 1 Then
                    Dim playerId As String
                    Dim refText As String
                    playerId = Trim(Left(line, eqPos - 1))
                    refText = Trim(Mid(line, eqPos + 1))
                    If playerId <> "" And IsNumeric(refText) Then
                        If map.Exists(playerId) Then
                            LogMessage(runId, "WARN", "Duplicate [Players] entry for '" & playerId & "' at line " & lineNo & " in " & iniPath & " (last value wins)")
                        End If
                        map(playerId) = CInt(refText)
                    Else
                        LogMessage(runId, "WARN", "Invalid [Players] entry at line " & lineNo & " in " & iniPath & ": " & line)
                    End If
                Else
                    LogMessage(runId, "WARN", "Invalid [Players] entry at line " & lineNo & " in " & iniPath & ": " & line)
                End If
            End If
        End If
    Loop

    ts.Close()
    On Error GoTo 0

    LoadPlayerMap = map
End Function

Function GetFeatureMapForParent(ByVal featureMapCache As Object, ByVal parentRef As Integer) As Object
    Dim cacheKey As String
    cacheKey = CStr(parentRef)

    If featureMapCache.Exists(cacheKey) Then
        GetFeatureMapForParent = featureMapCache(cacheKey)
        Exit Function
    End If

    Dim featureMap
    featureMap = CreateObject("Scripting.Dictionary")
    ResolveFeatureMap(parentRef, featureMap)
    featureMapCache.Add(cacheKey, featureMap)
    GetFeatureMapForParent = featureMap
End Function

Function IsIniTrue(ByVal rawValue As String) As Boolean
    Dim v As String
    v = LCase(Trim(rawValue))
    IsIniTrue = (v = "1" Or v = "true" Or v = "yes" Or v = "on")
End Function

Function ReadIniStringWithFallback(ByVal primarySection As String, ByVal primaryKey As String, ByVal primaryIni As String, ByVal fallbackSection As String, ByVal fallbackKey As String, ByVal fallbackIni As String, ByVal defaultValue As String) As String
    Dim marker As String
    marker = "__MISSING__"

    Dim primaryValue As String
    primaryValue = hs.GetINISetting(primarySection, primaryKey, marker, primaryIni)
    If primaryValue <> marker And Trim(primaryValue) <> "" Then
        ReadIniStringWithFallback = primaryValue
        Exit Function
    End If

    Dim fallbackValue As String
    fallbackValue = hs.GetINISetting(fallbackSection, fallbackKey, marker, fallbackIni)
    If fallbackValue <> marker And Trim(fallbackValue) <> "" Then
        ReadIniStringWithFallback = fallbackValue
    Else
        ReadIniStringWithFallback = defaultValue
    End If
End Function

Function ReadIniInt(ByVal section As String, ByVal key As String, ByVal defaultValue As Integer, ByVal iniFile As String) As Integer
    Dim raw As String
    raw = hs.GetINISetting(section, key, CStr(defaultValue), iniFile)
    If IsNumeric(raw) Then
        ReadIniInt = CInt(raw)
    Else
        ReadIniInt = defaultValue
    End If
End Function

Sub ResolveFeatureMap(ByVal parentRef As Integer, ByRef featureMap As Object)
    On Error Resume Next

    Dim parentDev
    parentDev = hs.GetDeviceByRef(parentRef)
    If parentDev Is Nothing Then Exit Sub

    Dim assocRefs
    assocRefs = parentDev.AssociatedDevices(hs)
    If Not IsArray(assocRefs) Then Exit Sub

    Dim childRef
    For Each childRef In assocRefs
        Dim childDev
        childDev = hs.GetDeviceByRef(CInt(childRef))
        If Not childDev Is Nothing Then
            Dim featureName As String
            featureName = LCase(CStr(childDev.Name(hs)))
            If featureName <> "" Then
                If Not featureMap.Exists(featureName) Then
                    featureMap.Add(featureName, CInt(childRef))
                End If
            End If
        End If
    Next

    On Error GoTo 0
End Sub

Sub SetFeatureString(ByVal featureMap As Object, ByVal featureName As String, ByVal value As String)
    Dim k As String
    k = LCase(featureName)
    If featureMap.Exists(k) Then
        SetDeviceStringIfChanged(CInt(featureMap(k)), value)
    End If
End Sub

Function FileNameFromPath(ByVal fullPath As String) As String
    If fullPath = "" Then
        FileNameFromPath = ""
        Exit Function
    End If

    Dim lastSlash As Integer
    Dim lastBackslash As Integer
    Dim sepPos As Integer
    lastSlash = InStrRev(fullPath, "/")
    lastBackslash = InStrRev(fullPath, "\")
    sepPos = lastSlash
    If lastBackslash > sepPos Then sepPos = lastBackslash

    If sepPos > 0 Then
        FileNameFromPath = Mid(fullPath, sepPos + 1)
    Else
        FileNameFromPath = fullPath
    End If
End Function

Function FormatDuration(ByVal milliseconds As Double) As String
    Dim totalSeconds As Long
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long

    If milliseconds < 0 Then milliseconds = 0

    totalSeconds = CLng(Fix(milliseconds / 1000))
    hours = totalSeconds \ 3600
    minutes = (totalSeconds Mod 3600) \ 60
    seconds = totalSeconds Mod 60

    FormatDuration = CStr(hours) & ":" & Right("0" & CStr(minutes), 2) & ":" & Right("0" & CStr(seconds), 2)
End Function

 ' BLPlex base formatter used by Duration/Progress/TimeRemaining
 ' (decompiled lines 13924-14162):
 ' milliseconds -> "<n> hrs, <n> mins, <n> secs" with only non-zero units.
Function FormatDurationBlPlex(ByVal milliseconds As Double) As String
    Dim totalSeconds As Long
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    Dim result As String

    If milliseconds < 0 Then milliseconds = 0

    totalSeconds = CLng(Fix(milliseconds / 1000))
    hours = totalSeconds \ 3600
    minutes = (totalSeconds Mod 3600) \ 60
    seconds = totalSeconds Mod 60

    result = ""
    If hours > 0 Then
        result = CStr(hours) & " hrs"
    End If

    If minutes > 0 Then
        If result <> "" Then result = result & ", "
        result = result & CStr(minutes) & " mins"
    End If

    If seconds > 0 Then
        If result <> "" Then result = result & ", "
        result = result & CStr(seconds) & " secs"
    End If

    FormatDurationBlPlex = result
End Function

' BLPlex Duration wrapper behavior (decompiled lines 13924-13986):
' if duration is non-numeric, pass through raw duration string.
Function ToBlPlexDurationText(ByVal durationRaw As String) As String
    Dim ms As Double
    ms = 0

    If IsNumeric(durationRaw) Then
        ms = CDbl(durationRaw)
        ToBlPlexDurationText = FormatDurationBlPlex(ms)
    Else
        ToBlPlexDurationText = durationRaw
    End If
End Function

' BLPlex Progress wrapper behavior (decompiled lines 13988-14068):
' zero -> "None"; non-numeric -> pass through raw viewOffset string.
Function ToBlPlexProgressText(ByVal viewOffsetRaw As String) As String
    Dim ms As Double
    Dim raw As String
    ms = 0
    raw = LCase(Trim(viewOffsetRaw))

    If IsNumeric(viewOffsetRaw) Then
        ms = CDbl(viewOffsetRaw)
        If ms > 0 Then
            ToBlPlexProgressText = FormatDurationBlPlex(ms)
        Else
            ToBlPlexProgressText = "None"
        End If
    ElseIf raw = "" Or raw = "unknown" Then
        ' Avoid transient "Unknown" at session start before first numeric offset arrives.
        ToBlPlexProgressText = "None"
    Else
        ToBlPlexProgressText = viewOffsetRaw
    End If
End Function

' BLPlex TimeRemaining wrapper behavior (decompiled lines 14070-14162):
' remaining > 0 -> formatted duration; remaining = 0 -> "Finished";
' remaining < 0 or invalid input -> "N/A".
Function ToBlPlexTimeRemainingText(ByVal durationRaw As String, ByVal viewOffsetRaw As String) As String
    Dim durationMs As Double
    Dim viewOffsetMs As Double
    Dim remainingMs As Double
    durationMs = 0
    viewOffsetMs = 0
    remainingMs = 0

    If IsNumeric(durationRaw) And IsNumeric(viewOffsetRaw) Then
        durationMs = CDbl(durationRaw)
        viewOffsetMs = CDbl(viewOffsetRaw)
        remainingMs = durationMs - viewOffsetMs
        If remainingMs > 0 Then
            ToBlPlexTimeRemainingText = FormatDurationBlPlex(remainingMs)
        ElseIf remainingMs = 0 Then
            ToBlPlexTimeRemainingText = "Finished"
        Else
            ToBlPlexTimeRemainingText = "N/A"
        End If
    Else
        ToBlPlexTimeRemainingText = "N/A"
    End If
End Function

Sub ApplyStoppedDefaults(ByVal debugEnabled As Boolean, ByVal runId As String, ByVal playerID As String, ByVal parentRef As Integer, ByVal featureMapCache As Object, ByVal showLastPlayedWhenStopped As Boolean)
    Dim featureMap As Object
    Dim notSetText As String
    featureMap = GetFeatureMapForParent(featureMapCache, parentRef)
    notSetText = "Not Set"

    DebugLog(debugEnabled, runId, "PlexPoll", "No active session for mapped player " & playerID & "; applying stopped defaults")

    SetFeatureString(featureMap, "PlayerIdentifier", playerID)
    If Not showLastPlayedWhenStopped Then
        SetFeatureString(featureMap, "Rating", notSetText)
        SetFeatureString(featureMap, "ContentRating", notSetText)
        SetFeatureString(featureMap, "Duration", notSetText)
        SetFeatureString(featureMap, "Progress", notSetText)
        SetFeatureString(featureMap, "TimeRemaining", notSetText)
        SetFeatureString(featureMap, "MediaType", notSetText)
        SetFeatureString(featureMap, "MediaTitle", notSetText)
        SetFeatureString(featureMap, "MediaFile", notSetText)
        SetFeatureString(featureMap, "MediaFilePath", notSetText)
        SetFeatureString(featureMap, "ArtworkThumb", notSetText)
        SetFeatureString(featureMap, "Lyrics", notSetText)
        SetFeatureString(featureMap, "Artist", notSetText)
        SetFeatureString(featureMap, "Album", notSetText)
        SetFeatureString(featureMap, "User", notSetText)
    End If

    SetParentStateByValue(parentRef, "Stopped")

    If featureMap.Exists("progress") And Not showLastPlayedWhenStopped Then
        SetDeviceValueIfChanged(CInt(featureMap("progress")), 0)
    End If
End Sub

Sub DebugLog(ByVal debugEnabled As Boolean, ByVal runId As String, ByVal logType As String, ByVal message As String)
    If debugEnabled Then
        hs.WriteLog(logType, "[" & runId & "] DEBUG: " & message)
    End If
End Sub

Sub LogMessage(ByVal runId As String, ByVal level As String, ByVal message As String)
    hs.WriteLog("PlexPoll", "[" & runId & "] " & level & ": " & message)
End Sub

Sub ValidateMappedFeatures(ByVal runId As String, ByVal playerMap As Object, ByVal featureMapCache As Object)
    Dim required
    required = Split("playeridentifier,rating,contentrating,duration,timeremaining,mediatype,mediatitle,mediafile,mediafilepath,artworkthumb,artist,album,user,progress", ",")

    Dim checkedParents
    checkedParents = CreateObject("Scripting.Dictionary")

    Dim playerId
    For Each playerId In playerMap.Keys
        Dim parentRef As Integer
        parentRef = CInt(playerMap(playerId))

        Dim parentKey As String
        parentKey = CStr(parentRef)
        If checkedParents.Exists(parentKey) Then
            GoTo NextParent
        End If
        checkedParents.Add(parentKey, True)

        Dim featureMap
        featureMap = GetFeatureMapForParent(featureMapCache, parentRef)
        Dim missing As String
        missing = ""

        Dim k
        For Each k In required
            If Not featureMap.Exists(CStr(k)) Then
                If missing <> "" Then missing = missing & ", "
                missing = missing & CStr(k)
            End If
        Next

        If missing <> "" Then
            LogMessage(runId, "WARN", "Parent ref " & parentRef & " missing feature(s): " & missing)
        End If
NextParent:
    Next
End Sub

Sub SetParentStateByValue(ByVal parentRef As Integer, ByVal stateLabel As String)
    ' Clear any stale custom device string so displayed status comes from VSP value map.
    SetDeviceStringIfChanged(parentRef, "")
    SetDeviceValueIfChanged(parentRef, ResolveParentStateValue(parentRef, stateLabel))
End Sub

Sub SetDeviceStringIfChanged(ByVal devRef As Integer, ByVal newValue As String)
    hs.SetDeviceString(devRef, newValue, True)
End Sub

Sub SetDeviceValueIfChanged(ByVal devRef As Integer, ByVal newValue As Double)
    hs.SetDeviceValueByRef(devRef, newValue, True)
End Sub

Function ResolveParentStateValue(ByVal parentRef As Integer, ByVal stateLabel As String) As Double
    On Error Resume Next

    Dim target As String
    target = LCase(stateLabel)

    Dim vspairs
    vspairs = hs.DeviceVSP_GetAllStatus(parentRef)
    If Not vspairs Is Nothing Then
        Dim p
        For Each p In vspairs
            Dim candidateValue As Double
            Dim candidateLabel As String
            candidateValue = CDbl(p.Value)
            candidateLabel = LCase(hs.DeviceVSP_GetStatus(parentRef, candidateValue, ePairStatusControl.Status))
            If candidateLabel = target Then
                ResolveParentStateValue = candidateValue
                On Error GoTo 0
                Exit Function
            End If
        Next
    End If

    ' Fallback map when no matching VSP label is found
    Select Case target
        Case "playing", "buffering"
            ResolveParentStateValue = 1
        Case "paused"
            ResolveParentStateValue = 2
        Case Else
            ResolveParentStateValue = 0
    End Select

    On Error GoTo 0
End Function
