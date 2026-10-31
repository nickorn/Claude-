Attribute VB_Name = "modSettings"
Option Explicit

' User settings in %APPDATA%\PatentTools\ (settings.json, blacklist.json,
' textblocks\). Missing files are created from the defaults embedded in
' modDefaults; keys missing in the user's file fall back to those defaults,
' so older settings files keep working after an update.

Private mUser As Object
Private mDefaults As Object
Private mLoaded As Boolean

Public Function ConfigFolder() As String
    ConfigFolder = Environ$("APPDATA") & "\PatentTools"
End Function

Public Function SettingsPath() As String
    SettingsPath = ConfigFolder() & "\settings.json"
End Function

Public Function BlacklistPath() As String
    BlacklistPath = ConfigFolder() & "\blacklist.json"
End Function

' Creates the config folder and default files on first run.
Public Sub EnsureUserConfig()
    Dim folder As String
    folder = ConfigFolder()
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    If Len(Dir$(folder & "\textblocks", vbDirectory)) = 0 Then MkDir folder & "\textblocks"
    If Len(Dir$(SettingsPath())) = 0 Then WriteUtf8File SettingsPath(), DecodeJsonUnicode(DefaultSettingsJson())
    If Len(Dir$(BlacklistPath())) = 0 Then WriteUtf8File BlacklistPath(), DecodeJsonUnicode(DefaultBlacklistJson())
End Sub

' Forces settings.json to be read again on the next GetSettingValue call.
' Every ribbon command calls this first, so edits apply without restarting Word.
Public Sub ReloadSettings()
    mLoaded = False
    Set mUser = Nothing
End Sub

' Returns a scalar setting by dotted path, e.g. "renumber.trackChangesMode".
Public Function GetSettingValue(ByVal path As String, ByVal fallback As Variant) As Variant
    Dim v As Variant
    If Not mLoaded Then LoadSettings
    v = JsonLookup(mUser, path)
    If IsEmpty(v) Then v = JsonLookup(mDefaults, path)
    If IsEmpty(v) Or IsNull(v) Then v = fallback
    GetSettingValue = v
End Function

Public Sub OpenSettingsFile()
    On Error GoTo Fail
    EnsureUserConfig
    Shell "notepad.exe """ & SettingsPath() & """", vbNormalFocus
    ReloadSettings
    Exit Sub
Fail:
    MsgBox MsgConfigFolderError(ConfigFolder(), Err.Description), vbExclamation, AppTitle()
End Sub

Private Sub LoadSettings()
    Dim errText As String
    If mDefaults Is Nothing Then Set mDefaults = JsonParse(DecodeJsonUnicode(DefaultSettingsJson()))
    Set mUser = Nothing

    On Error Resume Next
    EnsureUserConfig
    If Err.Number <> 0 Then errText = MsgConfigFolderError(ConfigFolder(), Err.Description)
    Err.Clear
    If Len(errText) = 0 And Len(Dir$(SettingsPath())) > 0 Then
        Set mUser = JsonParse(ReadUtf8File(SettingsPath()))
        If Err.Number <> 0 Then
            errText = MsgSettingsError(SettingsPath(), Err.Description)
            Set mUser = Nothing
        End If
    End If
    On Error GoTo 0

    mLoaded = True
    If Len(errText) > 0 Then MsgBox errText, vbExclamation, AppTitle()
End Sub

' Terms of blacklist.json for "en" or "de". Falls back to the embedded default
' list if the user's file cannot be read.
Public Function GetBlacklistTerms(ByVal lang As String) As Collection
    Dim root As Object, terms As New Collection, v As Variant, errText As String
    On Error Resume Next
    EnsureUserConfig
    Err.Clear
    Set root = JsonParse(ReadUtf8File(BlacklistPath()))
    If Err.Number <> 0 Then
        errText = MsgSettingsError(BlacklistPath(), Err.Description)
        Set root = Nothing
    End If
    Err.Clear
    On Error GoTo 0
    If Len(errText) > 0 Then MsgBox errText, vbExclamation, AppTitle()
    If root Is Nothing Then Set root = JsonParse(DecodeJsonUnicode(DefaultBlacklistJson()))
    If root.Exists(lang) Then
        For Each v In root.Item(lang)
            If Not IsObject(v) Then terms.Add CStr(v)
        Next v
    End If
    Set GetBlacklistTerms = terms
End Function

' ---------------------------------------------------------------- UTF-8 files

Public Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText(-1)
    st.Close
End Function

Public Sub WriteUtf8File(ByVal path As String, ByVal fileText As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText fileText
    st.SaveToFile path, 2 ' overwrite
    st.Close
End Sub
