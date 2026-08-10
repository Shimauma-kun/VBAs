Attribute VB_Name = "mdlConvertPBrush"
Sub ConvertPBrush()
    Dim imgFolder As String
    imgFolder = Environ("USERPROFILE") & "\Desktop\image"

    If Dir(imgFolder, vbDirectory) = "" Then
        MkDir imgFolder
    End If

    ' --- ① 先にPBrushが実際にあるか確認 ---
    Dim targets As Collection
    Set targets = New Collection
    Dim i As Long, s As Shape
    For i = 1 To ActiveDocument.Shapes.Count
        Set s = ActiveDocument.Shapes(i)
        If s.Type = msoEmbeddedOLEObject Then
            If s.OLEFormat.ClassType = "PBrush" Then
                targets.Add s.Name
            End If
        End If
    Next i

    If targets.Count = 0 Then
        MsgBox "PBrushオブジェクトは見つかりませんでした。"
        Exit Sub
    End If

    ActiveDocument.Save

    ' --- ② docmをzipとしてコピーし、丸ごと展開する ---
    Dim tempRoot As String
    tempRoot = Environ("TEMP") & "\pbrush_extract_" & Format(Now, "yyyymmddhhnnss")
    MkDir tempRoot

    Dim zipPath As String
    zipPath = tempRoot & "\src.zip"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile ActiveDocument.FullName, zipPath, True

    Dim extractFolder As String
    extractFolder = tempRoot & "\extracted"
    MkDir extractFolder

    ExtractZipAll zipPath, extractFolder

    Dim docXmlPath As String, relsPath As String
    docXmlPath = extractFolder & "\word\document.xml"
    relsPath = extractFolder & "\word\_rels\document.xml.rels"

    If Dir(docXmlPath) = "" Or Dir(relsPath) = "" Then
        MsgBox "展開に失敗しました(document.xmlが見つかりません)。"
        Exit Sub
    End If

    Dim docXml As String, relsXml As String
    docXml = ReadTextUTF8(docXmlPath)
    relsXml = ReadTextUTF8(relsPath)

    ' --- ③ PBrushのプレビュー画像(v:imagedata)を、文書内の出現順に洗い出す ---
    Dim reObj As Object
    Set reObj = CreateObject("VBScript.RegExp")
    reObj.Global = True
    reObj.IgnoreCase = True
    reObj.pattern = "<v:imagedata[^>]*r:id=""(rId\d+)""[^>]*/>[\s\S]{0,2000}?" & _
                     "<o:OLEObject[^>]*ProgID=""PBrush""[^>]*r:id=""(rId\d+)""[^>]*/>"

    Dim mats As Object
    Set mats = reObj.Execute(docXml)

    Debug.Print "PBrush検出(XML解析): " & mats.Count & " 件"

    If mats.Count <> targets.Count Then
        MsgBox "警告: シェイプから数えたPBrush数(" & targets.Count & ")と、" & _
               "XMLから検出した数(" & mats.Count & ")が一致しません。" & vbCrLf & _
               "処理を中止します。"
        Exit Sub
    End If

    ' --- ④ 各プレビューのrIdから実ファイル名を割り出し、imageフォルダへコピー ---
    Dim previewPaths() As String
    ReDim previewPaths(1 To mats.Count)

    Dim m As Object, idx As Long
    idx = 0
    For Each m In mats
        idx = idx + 1
        Dim previewRid As String
        previewRid = m.SubMatches(0)

        Dim relPattern As String
        relPattern = "<Relationship[^>]*Id=""" & previewRid & """[^>]*/>"

        Dim reRel As Object
        Set reRel = CreateObject("VBScript.RegExp")
        reRel.Global = False
        reRel.IgnoreCase = True
        reRel.pattern = relPattern

        Dim relMatch As Object
        Set relMatch = reRel.Execute(relsXml)

        If relMatch.Count = 0 Then
            MsgBox "rels内に " & previewRid & " が見つかりません。"
            Exit Sub
        End If

        Dim relWhole As String
        relWhole = relMatch(0).Value

        Dim reTarget As Object
        Set reTarget = CreateObject("VBScript.RegExp")
        reTarget.Global = False
        reTarget.pattern = "Target=""media/([^""]+)"""
        Dim targetMatch As Object
        Set targetMatch = reTarget.Execute(relWhole)

        If targetMatch.Count = 0 Then
            MsgBox previewRid & " のTargetが取得できません。"
            Exit Sub
        End If

        Dim mediaFn As String
        mediaFn = targetMatch(0).SubMatches(0)

        Dim mediaSrc As String
        mediaSrc = extractFolder & "\word\media\" & mediaFn

        Dim ext As String
        ext = Mid(mediaFn, InStrRev(mediaFn, "."))

        Dim outName As String
        outName = "pbrush_preview_" & Format(idx, "00") & ext
        Dim outPath As String
        outPath = imgFolder & "\" & outName

        fso.CopyFile mediaSrc, outPath, True
        previewPaths(idx) = outPath

        Debug.Print "  [" & idx & "] media/" & mediaFn & " -> " & outName
    Next m

    ' --- ⑤ 抽出した画像を、PBrushシェイプに順番に割り当てる ---
    Dim nm As Variant, counter As Long
    counter = 0
    For Each nm In targets
        counter = counter + 1

        Dim srcPath As String
        srcPath = previewPaths(counter)

        Dim shp As Shape
        Set shp = ActiveDocument.Shapes(CStr(nm))

        Dim leftPos As Single, topPos As Single, w As Single, h As Single
        Dim wrapType As Long, anchorRange As Range
        leftPos = shp.Left
        topPos = shp.Top
        w = shp.Width
        h = shp.Height
        wrapType = shp.WrapFormat.Type
        Set anchorRange = shp.Anchor

        shp.Delete

        Dim newShp As Shape
        Set newShp = ActiveDocument.Shapes.AddPicture( _
            FileName:=srcPath, LinkToFile:=False, SaveWithDocument:=True, _
            Left:=leftPos, Top:=topPos, Width:=w, Height:=h, Anchor:=anchorRange)
        newShp.WrapFormat.Type = wrapType
        newShp.Name = CStr(nm)

        Debug.Print "変換完了(" & counter & "): " & nm & " <- " & srcPath & " (埋め込み)"
    Next nm

    ' --- 後片付け ---
    On Error Resume Next
    Kill extractFolder & "\word\media\*.*"
    Kill zipPath
    On Error GoTo 0

    MsgBox counter & "件のPBrushオブジェクトを埋め込み画像に変換しました。" & vbCrLf & _
           "見た目が正しいか、必ず目視で確認してください。"
End Sub

Sub ExtractZipAll(zipPath As String, destFolder As String)
    If Dir(zipPath) = "" Then
        MsgBox "zipファイルが見つかりません: " & zipPath
        End
    End If

    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")

    Dim zipFolder As Object
    Set zipFolder = shellApp.Namespace(CVar(zipPath))

    If zipFolder Is Nothing Then
        MsgBox "zipFolderの取得に失敗しました。パス: " & zipPath
        End
    End If

    Dim destNs As Object
    Set destNs = shellApp.Namespace(CVar(destFolder))

    If destNs Is Nothing Then
        MsgBox "destNsの取得に失敗しました。パス: " & destFolder
        End
    End If

    destNs.CopyHere zipFolder.Items, 4 + 16   ' 4=進捗ダイアログ非表示, 16=全て上書きしてよい

    ' 非同期なので、目印のファイルが出現するまで待つ(Wordには Application.Wait が無いため Timer で代用)
    Dim startTime As Single
    startTime = Timer
    Do While Dir(destFolder & "\word\document.xml") = ""
        DoEvents
        If Timer - startTime > 60 Then
            Exit Do   ' 60秒でタイムアウト
        End If
    Loop
End Sub


' --- UTF-8のテキストファイルを文字列として読み込む ---
Function ReadTextUTF8(path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' テキスト
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadTextUTF8 = st.ReadText
    st.Close
End Function


