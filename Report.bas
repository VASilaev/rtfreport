﻿Attribute VB_Name = "Report"
 Option Compare Text
 Option Explicit
 
 Public Enum tOperationType
  opEQ = 1 ' равно
  opNEQ = 2 ' не равно
  opGR = 4 ' больше
  opLS = 8 ' меньше
  opNLS = 16 ' не меньше
  opNGR = 32 ' не больше
  opIN = 128 ' в списке
  opNIN = 256 ' не в списке
  opcont = 512 ' содержит
  opSTART = 1024 ' начинается
  opBTW = 2048 ' между
  opBTWWL = 6144 ' между без левого
  opBTWWR = 10240 ' между без правого
  opBTWWB = 14336 ' между без обоих
  opNCont = 32768 ' не содержит
End Enum
  
Public Sub InstallRepSystem()
'Создает необходимые таблицы. Для хранения шаблонов внутри таблицы
  If Not IsHasRepTable() Then
    With CurrentDb()
      .Execute "CREATE TABLE t_rep " _
        & "(id counter CONSTRAINT PK_rep PRIMARY KEY, " _
        & "sCaption CHAR(255), sOrignTemplate memo, " _
        & "dEditTemplate date, sDescription char(255),clTemplate memo);"
      .TableDefs.Refresh
    End With
  End If
End Sub

Public Sub InstallReportTemplate()
'Добавляет файл в хранилище шаблонов, после чего сам файл можно удалить, а шаблон вызывать по коду или имени
  InstallRepSystem

  Dim dlgOpenFile, sFileName As String, idReport, sFilePath As String, atmp
  Set dlgOpenFile = Application.FileDialog(3)
  With dlgOpenFile
    .Filters.Clear
    .InitialFileName = CurrentProject.Path
    .Filters.Add "RTF шаблон", "*.rtf", 1
    .AllowMultiSelect = False
    .Title = "Выберите шаблон"
    If (.Show = -1) And (.SelectedItems.Count > 0) Then
      sFilePath = .SelectedItems(1)
    End If
  End With
  Set dlgOpenFile = Nothing
  
  sFileName = GetFile(sFilePath)
  sFileName = Mid(sFileName, 1, Len(sFileName) - Len(GetExt(sFileName)))
  
  idReport = SelectOneValue("select id from t_rep where ucase(sCaption) = '" & sFileName & "'")
  If IsEmpty(idReport) Then
    If UCase(Left(sFilePath, Len(CurrentProject.Path))) = UCase(CurrentProject.Path) Then sFilePath = "." & Mid(sFilePath, Len(CurrentProject.Path) + 1)
    CurrentDb().Execute "insert into t_rep (sCaption, sOrignTemplate) values ('" & Replace(sFileName, "'", "''") & "','" & Replace(sFilePath, "'", "''") & "');"
    idReport = SelectOneValue("select id from t_rep where ucase(sCaption) = '" & UCase(sFileName) & "'")
    atmp = GetTemplate(CLng(idReport))
    MsgBox "Отчет с именем """ & sFileName & """ зарегистрирован с кодом " & idReport
  Else
    MsgBox "Отчет с именем """ & sFileName & """ уже существует по кодом " & idReport
  End If
End Sub



Public Sub PrintReport(vReport, Optional ByRef dic As Object, Optional sFile As String = "")
'Запускает формирование документа из шаблона
'#param vReport: Идентификатор шаблона.
'Если число, то ищется в таблице t_rep, в противном случае считается что это имя файла.
'Для поиска относительно местоположения БД используйте в начале `.\`.
'Если такого файла не существует, то шаблон ищется по заголовку (`sCaption`) в таблице t_rep.
'#param dic: Словарь с окружением, можно передать nothing если явных входных параметров нет
'#param sFile: Имя выходного файла, если его не указать то будет создан во временной папке с именем tmp_n где n - порядковый номер

 Dim fso
 Dim tf 'As TextStream
 Dim asTemplate, tValue, i As Integer
 Dim sPathOrig As String, sExtension As String

 Set fso = CreateObject("scripting.FileSystemObject")
 
 If dic Is Nothing Then
  Set dic = CreateObject("Scripting.Dictionary")
  dic.CompareMode = 1
 End If

 If IsNumeric(vReport) Then
   If IsHasRepTable() Then
     asTemplate = Report.GetTemplate(CLng(vReport))
   Else
     Err.Raise 1000, , "Не найден шаблон """ & vReport & """"
   End If
 Else
  
  sPathOrig = vReport
  If Left(sPathOrig, 2) = ".\" Then sPathOrig = CurrentProject.Path & Mid(sPathOrig, 2)
   
  sExtension = LCase(fso.GetExtensionName(sPathOrig))
   
  If fso.FileExists(sPathOrig) Then
    Select Case sExtension
      Case "rtf"
        asTemplate = Array("rtf", PrepareRTF(sPathOrig))
     End Select
  Else
    i = Empty
    If IsHasRepTable() Then i = SelectOneValue("select id from t_rep where ucase(sCaption) = '" & UCase(vReport) & "'")
    If IsEmpty(i) Then
      Err.Raise 1000, , "Не найден шаблон """ & vReport & """"
    Else
      asTemplate = Report.GetTemplate(CLng(i))
    End If
  End If
 End If
 
 dic("extension") = asTemplate(0)
 dic("Date") = Date
 dic("Now") = Now
  
 If sFile = "" Then
   On Error Resume Next
   For i = 0 To 1000
    sFile = fso.GetSpecialFolder(2) & "\tmp_" & i & "." & asTemplate(0)
    If fso.FileExists(sFile) Then
     fso.DeleteFile sFile, True
     If Not fso.FileExists(sFile) Then Exit For
    Else
     Exit For
    End If
   Next
 ElseIf fso.FileExists(sFile) Then
  Dim sBaseName As String
   sBaseName = fso.GetBaseName(sFile) & "_"
   sExtension = "." & fso.GetExtensionName(sFile)
   For i = 0 To 1000
    sFile = sBaseName & i & sExtension
    If Not fso.FileExists(sFile) Then
     Exit For
    End If
   Next
 End If
 
 On Error GoTo 0
 Set tf = fso.CreateTextFile(sFile)
 Dim svTemplate As String
 svTemplate = asTemplate(1)
 Report.MakeReport svTemplate, tf, dic
 tf.Close
 Set tf = Nothing
 Set fso = Nothing
 
 If InSet(asTemplate(0), "rtf") Then Shell "winword """ & sFile & """", vbNormalFocus
 If InSet(asTemplate(0), "txt") Then Shell "notepad """ & sFile & """", vbNormalFocus
 asTemplate = Array()
 
End Sub

Private Function IsHasRepTable()
'Проверка наличия в БД таблицы `t_rep`

  Dim vTbl, vFld, vDB
  IsHasRepTable = True
  Set vDB = CurrentDb()
  On Error GoTo onCreate
  Set vTbl = vDB.TableDefs("t_rep")
  On Error GoTo 0
  Exit Function
onCreate:
  On Error GoTo 0
  IsHasRepTable = False
End Function




Function BuildParam(pDic, ParamArray pdata())
'Обновляет в контексте переменные
'`BuildParam(pDic, Key, Value [, Key, Value])`
'#param pDic: Текст
'#param Key: Имя переменной
'#param Value: Значение переменной. Объекты должны быть завернуты в массив: `array(MyObject)`

  Dim i, tmp, vData
  vData = pdata
  If UBound(vData) = 1 Then
    If IsNull(vData(0)) And IsArray(vData(1)) Then vData = vData(1)
  End If
  
  If pDic Is Nothing Then
    Set pDic = CreateObject("Scripting.Dictionary")
    pDic.CompareMode = 1
  End If
 
  i = LBound(vData)
  Do While i <= UBound(vData)
    If IsArray(vData) Then
      BuildParam pDic, Null, vData(i)
      i = i + 1
    ElseIf i < UBound(vData) Then
      If IsNull(vData(0)) And IsArray(vData(1)) Then
        BuildParam pDic, Null, vData(i + 1)
      Else
        pDic(vData(i) & "") = vData(i + 1)
      End If
      i = i + 2
    Else
      Err.Raise 1000, , "Не парное число параметров!"
    End If
    
  Loop
  Set BuildParam = pDic
End Function


Public Function GetLocation(ByRef spText As String, npPos As Long) As String
'По абсолютной позиции в тексте возвращет сообщение в формате "Строка <номер_строки>:<номер_символа_в_строке>"
'#param spText: Текст
'#param npPos: Абсолютная позиция от начала строки

 Dim nvLine As Long, nvColumn As Long, nvCrPos As Long, nvtCrPos As Long
 nvtCrPos = 1
 nvLine = 0
 
 Do
  nvLine = nvLine + 1
  nvCrPos = nvtCrPos
  nvtCrPos = InStr(1, spText, vbCr)
 Loop While nvtCrPos <> 0 And nvtCrPos < npPos
 
 nvColumn = npPos - nvCrPos
 
 If Mid(spText, nvCrPos + 1, 1) = vbLf Then nvColumn = nvColumn - 1
 
 GetLocation = "Строка " & nvLine & ":" & nvColumn
End Function

Function LPad(s As String, ch As String, TotalCnt As Integer) As String
'Дополняет текст слева символами для достижения заданной длины. Если текст изначально длиньше, то он усекается.
'#param s: Исходный текст
'#param ch: Добавляемый паттерн
'#param TotalCnt: Максимальная длина текста

 Dim t As Double
 t = (TotalCnt - Len(s)) \ Len(ch)
 If t - Int(t) > 0 Then t = Int(t) + 1
 LPad = s
 Do While t > 0
  LPad = ch & LPad
  t = t - 1
 Loop
 LPad = Right(LPad, TotalCnt)
End Function
 
 
 
Private Function GetEscape(ByRef sBuf As String, ByRef iPOS As Long) As String
'Внутренняя функция разбора RTF. Возвращает экрнаированное значение за символом `\`
'#param sBuf: Буфер
'#param iPOS: Текущая позиция в буфере. Позиция обновляется.

  If Mid(sBuf, iPOS, 1) = "\" Then
   Select Case Mid(sBuf, iPOS + 1, 1) '= "\"
    Case "'"
     GetEscape = Chr(CInt("&h" & Mid(sBuf, iPOS + 2, 2)))
     'iPOS = iPOS + 4
    Case "\", "{", "}"
     GetEscape = Mid(sBuf, iPOS + 1, 1)
     'iPOS = iPOS + 2
    Case Else
     GetEscape = ""
   End Select
  Else
   GetEscape = ""
  End If
 End Function
 
Private Function GetToken(ByRef sBuf As String, ByRef iPOS As Long)
'Внутренняя функция разбора RTF. Разбирает текущую лексему в тексте.
'Возвращет ее значение в следующем формате: <Код_Лексемы><Значение>.
'Код лексем принимают следующие значение:
' {*} с - Управляющая конструкция. Если достигнут конец файла, то возвращается `cEOF`
' {*} t - Текст
'#param sBuf: Буфер
'#param iPOS: Текущая позиция в буфере. Позиция обновляется.

  Dim State As Integer
  Dim LStr As String, ch As String
  
  '0, 10 - start

  '50 - text
  '100 - control
  
  GetToken = ""
  State = 0
  
  Do While True
   If iPOS > Len(sBuf) Then
    GetToken = "cEOF"
    Exit Do
   End If
   
   Select Case Mid(sBuf, iPOS, 1)
    Case Chr(13), Chr(10)
     iPOS = iPOS + 1
    Case " "
     Select Case State
      Case 0
       State = 10
      Case 10
       GetToken = GetToken & " "
       State = 50
      Case 50
       GetToken = GetToken & " "
      Case 100
       'iPOS = iPOS + 1
       GetToken = "c" & GetToken
       
       Exit Do
     End Select
     iPOS = iPOS + 1
    Case "{"
     If State = 0 Or State = 10 Then
      GetToken = "{"
      iPOS = iPOS + 1
      State = 100
     End If
     GetToken = IIf(State = 50, "t", "c") & GetToken
     Exit Do
    Case "}"
     If State = 0 Or State = 10 Then
      GetToken = "}"
      iPOS = iPOS + 1
      State = 100
     End If
     GetToken = IIf(State = 50, "t", "c") & GetToken
     Exit Do
    Case "\"
     ch = GetEscape(sBuf, iPOS)
     If ch = "" Then
      Select Case State
       Case 0, 10
        GetToken = "\"
        State = 100
        iPOS = iPOS + 1
       Case 50
        GetToken = "t" & GetToken
        Exit Do
       Case Else
        GetToken = "c" & GetToken
        Exit Do
      End Select
     Else
      Select Case State
       Case 0, 10
        GetToken = ch
        State = 50
        iPOS = iPOS + IIf(ch = "\" Or ch = "{" Or ch = "}", 2, 4)
       Case 50
        GetToken = GetToken & ch
        iPOS = iPOS + IIf(ch = "\" Or ch = "{" Or ch = "}", 2, 4)
       Case Else
        GetToken = "c" & GetToken
        Exit Do
      End Select
     End If
   Case Else
    If State = 0 Or State = 10 Then State = 50
    GetToken = GetToken & Mid(sBuf, iPOS, 1)
    iPOS = iPOS + 1
   End Select
  Loop

 End Function


 
Private Function SkipBlock(ByRef sBuf As String, ByRef iPOS As Long)
'Внутренняя функция разбора RTF. Разбирает текст до символа `}`. Вложенные теги так же пропускаются. Возвращает текст из пропущенного блока.
'#param sBuf: Буфер
'#param iPOS: Текущая позиция в буфере. Позиция обновляется.


  Dim ch As String
  SkipBlock = ""
  ch = GetToken(sBuf, iPOS)
  Do While ch <> "c}" And ch <> "cEOF"
   If ch = "c{" Then
    SkipBlock = SkipBlock & SkipBlock(sBuf, iPOS)
   End If
   ch = GetToken(sBuf, iPOS)
   If Left(ch, 1) = "t" Then
     SkipBlock = SkipBlock & ch
   End If
  Loop

 End Function


Private Function ParseField(ByRef sBuf As String, ByRef iPOS As Long)
'Внутренняя функция разбора RTF. Разбирает поле и возвращает массив [<Видимый_текст>,<Формат_текста>]. Формат берется на начало строки.
'#param sBuf: Буфер
'#param iPOS: Текущая позиция в буфере. Позиция обновляется.


  Dim sOpt As String
  Dim CP As Integer
  Dim State As Integer, ch As String, txt As String, sTmpToken As String
  
  
  State = 0
  ch = GetToken(sBuf, iPOS)
  
  Do While ch <> "cEOF"
  Select Case State
   Case 0
    Select Case LCase(ch)
     Case "c\flddirty", "c\fldedit", "c\fldlock", "c\fldpriv"
     Case "c{"
      State = 1 ' ждем \*
    End Select
   Case 1
    If ch = "c\*" Then
     State = 2 'ждем \fldinst
    Else
     Err.Raise 2001
    End If
   Case 2
    If ch = "c\fldinst" Then
     State = 3 ' ждем первого кортежа с текстом
    Else
     Debug.Print iPOS
     Err.Raise 1002
    End If
   Case 3, 13
    
    Select Case LCase(ch)
     Case "tref"
     Case "c{"
      ch = GetToken(sBuf, iPOS)
      If State = 3 Then sOpt = ""
      Do While ch <> "c}" And ch <> "cEnd"
       If ch = "c\line" Or ch = "c\par" Then
        txt = txt & vbCrLf
       ElseIf Left(ch, 1) = "c" Then
        If ch = "c{" Then SkipBlock sBuf, iPOS
        If State = 3 Then
          sTmpToken = Mid(ch, 2)
          'не все символы форматирования попадут
          If Not (Mid(sTmpToken, 1, 5) = "\lang" Or _
                  sTmpToken = "\rtlch" Or _
                  sTmpToken = "\ltrch" _
                 ) Then
            sOpt = sOpt & sTmpToken
          End If
        End If
       Else
        txt = txt & Mid(ch, 2)
        State = 13 'параметры берутся только из первого заполненого картежа
       End If
       ch = GetToken(sBuf, iPOS)
      Loop
     Case "c}"
      State = 20
      ch = GetToken(sBuf, iPOS)
      If ch = "c{" Then
       SkipBlock sBuf, iPOS
      Else
       Debug.Print iPOS
       Err.Raise 1001
      End If
      ch = GetToken(sBuf, iPOS)
      If ch = "c}" Then
       ParseField = Array(txt, sOpt)
       Exit Do
      Else
       Debug.Print iPOS
       Err.Raise 1001
      End If
    End Select
  End Select
  ch = GetToken(sBuf, iPOS)
  Loop
 End Function
 


Private Function FindParForSkip(ByRef sBuf As String, ByRef iPOS As Long) As String
'Внутренняя функция разбора RTF. Удаляет все теги с их внутренним содержимым из буфера
'#param rtf: Буфер
'#param tag: Удаляемый тег

  Dim ch As String
  ch = GetToken(sBuf, iPOS)
  Do While ch <> "c}" And ch <> "cEOF"
   If ch = "c\par" Or Left(ch, 1) = "t" Or ch = "ceof" Or ch = "c\field" Then
     FindParForSkip = ch
     Exit Function
   End If
   
   If ch = "c{" Then
    ch = LCase(FindParForSkip(sBuf, iPOS))
   Else
    ch = GetToken(sBuf, iPOS)
   End If

  Loop
  FindParForSkip = ""
End Function

Function RemoveTag(rtf, tag) As String
'Внутренняя функция разбора RTF. Удаляет все теги с их внутренним содержимым из буфера
'#param rtf: Буфер
'#param tag: Удаляемый тег

  Dim tp As Long, sSpace As String
  
  RemoveTag = rtf
  tp = InStr(1, RemoveTag, tag, vbTextCompare)
  Do While tp > 0
   Dim ep As Long
   
   ep = tp + 1
   SkipBlock RemoveTag, ep
 
   If Mid(rtf, tp - 1, 1) <> " " And Not InSet(Mid(RemoveTag, ep), " ", "\", "{", "}") Then sSpace = " " Else sSpace = ""

   RemoveTag = Mid(RemoveTag, 1, tp - 1) & sSpace & Mid(RemoveTag, ep)
   tp = InStr(1, RemoveTag, tag, vbTextCompare)
  Loop

End Function

Private Function RTF_SkipPar(ByRef sBuf As String, ByRef tp As Long) As String
'Внутренняя функция разбора RTF. Пропускает перевод строки если он идет непосредственно за текстом.
'#param sBuf: Буфер
'#param tp: Текущая позиция разбора

  Dim LP As Long, sFnc As String
  LP = tp

  Do While True

   If RTF_SkipPar = "c{" Then
    RTF_SkipPar = LCase(RTF_SkipPar(sBuf, tp))
   Else
    RTF_SkipPar = LCase(GetToken(sBuf, tp))
   End If

   If RTF_SkipPar = "c\par" Then Exit Do

   If Left(RTF_SkipPar, 1) = "t" Or _
      InSet(RTF_SkipPar, "c\field", "c\cell", "c\sect", "ceof") Then
    tp = LP
    Exit Do
   End If
  Loop

End Function


Private Function InsertAddress(ts As String, adr As Long, iPOS As Long) As String
'Внутрення функция формирования шаблона. Вставляет адрес (8 символов в шестнадцетиричном виде) вместо заглушки
'#param ts: Текущий буфер
'#param adr: Вставляемый адрес
'#param iPOS: Смещение в буфере куда нужно вставить адрес

  InsertAddress = Mid(ts, 1, iPOS) & LPad(Hex(adr), "0", 8) & Mid(ts, iPOS + 9)
End Function

Public Function PrepareRTF(sFile As String) As String
'Компилирует RTF файл в внутренний формат шаблона
'#param sFile: Содержимое RTF файла


 Dim fso
 Dim tf
 Dim ts As String
 Dim CP As Long, tp As Long, PosO As Long, PosC As Long
 Dim iScanCnt
 Dim iStrucLevel As Integer
 Dim aStrucStack(128) As Variant
 Dim nState, skipConst '/**Флаг пропуска куска RTF, в шаблон не включается*/
 Dim Res As String
 Dim MetList As String
 Dim sFmtTmp, sTXT, SFMT, sFnc, sOpt
 Dim adr

 Set fso = CreateObject("scripting.FileSystemObject")
  
 iStrucLevel = 0
 iScanCnt = 0
 
 Set tf = fso.OpenTextFile(sFile)
 ts = tf.ReadAll
 'ts = Replace(Replace(ts, Chr(13), ""), Chr(10), "")
 tf.Close
 
 CP = 1
 Res = "GOTO00000015        " 'если будет список меток то заменится на call
 
 Dim re
 Set re = CreateObject("VBScript.RegExp")
 re.IgnoreCase = True
 re.Global = True
 
'Удалим историю редактирования
 re.Pattern = "( |\r\n)?\\pararsid\d+"
 ts = re.Replace(ts, "")
 re.Pattern = "( |\r\n)?\\insrsid\d+"
 ts = re.Replace(ts, "")
 re.Pattern = "( |\r\n)?\\charrsid\d+"
 ts = re.Replace(ts, "")
 re.Pattern = "( |\r\n)?\\sectrsid\d+"
 ts = re.Replace(ts, "")
 re.Pattern = "( |\r\n)?\\styrsid\d+"
 ts = re.Replace(ts, "")
 re.Pattern = "( |\r\n)?\\tblrsid\d+"
 ts = re.Replace(ts, "")
 
'Удалим лишние плюшки WORD
 ts = RemoveTag(ts, "{\*\themedata")
 ts = RemoveTag(ts, "{\*\colorschememapping")
 ts = RemoveTag(ts, "{\*\latentstyles")
 ts = RemoveTag(ts, "{\*\datastore")
 ts = RemoveTag(ts, "{\*\rsidtbl")
 ts = RemoveTag(ts, "{\*\xmlnstbl")
 ts = RemoveTag(ts, "{\*\panose")
 ts = RemoveTag(ts, "{\*\blipuid")
 ts = RemoveTag(ts, "{\nonshppict")
 ts = RemoveTag(ts, "{\sp{\sn metroBlob")


' Set tf = fso.CreateTextFile(IIf(fso.GetParentFolderName(sFile) <> "", fso.GetParentFolderName(sFile) & "\", "") & fso.GetBaseName(sFile) & "L.rtf", True)
' tf.Write ts
' tf.Close
 
 re.Multiline = True
 re.IgnoreCase = True
 re.Global = False
 re.Pattern = "^\s*[_0-9а-яa-zё]+\s*\(.*\)$"
 
 nState = 0

 iStrucLevel = 1
 aStrucStack(iStrucLevel - 1) = Array(-1)
 
  skipConst = False

 
 Do While CP <= Len(ts)
 
  Select Case nState
  Case 0

  ' ищем данные до следующего Field
  tp = InStr(CP, ts, "{\field", vbTextCompare)
  
  If tp <> 0 Then
   ' у нас есть поле

   If Trim(Mid(ts, CP, tp - CP)) <> "" And Not skipConst Then Res = Res & "PRNT" & LPad(Hex(tp - CP), "0", 8) & Mid(ts, CP, tp - CP)
   CP = tp
   tp = tp + 7
   sFmtTmp = ParseField(ts, tp)

   sTXT = Trim(IIf(LCase(Left(Trim(sFmtTmp(0)), 3)) = "ref", Mid(Trim(sFmtTmp(0)), 4), sFmtTmp(0)))
   SFMT = sFmtTmp(1)

   re.Pattern = "^\s*[_0-9а-яa-zё]+\s*\(.*\)$"
   
   If Not re.test(Replace(sTXT, vbCrLf, "")) Then
    'это не правильное поле оставляем его как есть
    If Trim(Mid(ts, CP, tp - CP)) <> "" Then Res = Res & "PRNT" & LPad(Hex(tp - CP), "0", 8) & Mid(ts, CP, tp - CP)
   Else
     sFnc = Mid(sTXT, 1, InStr(1, sTXT, "(") - 1)
     sOpt = Mid(sTXT, InStr(1, sTXT, "(") + 1)
     'If Right(sOpt, 1) <> ")" Then MsgBox "Ожидается в конце ')' в выражении " & stxt
     sOpt = Trim(Left(sOpt, Len(sOpt) - 1))


     nState = 2
     skipConst = False
    
    Select Case LCase(Trim(sFnc))
     Case "scan"
      If sOpt = "" Then Err.Raise 1020, , "Ожидается выражение после scan"
      
      Dim nForPos, svCursorName, svSQLBody: nForPos = InStr(LCase(sOpt), " for ")
      If nForPos = 0 Then Err.Raise 1021, , "Параметр scan должен быть вида <Имя курсора> for <Выражение строкового вида>"
       
       
      svCursorName = Mid(sOpt, 1, nForPos - 1)
      svSQLBody = Mid(sOpt, nForPos + 5)
      
      Res = Res & "OPRS" & LPad(Hex(Len(svCursorName)), "0", 3) & svCursorName & LPad(Hex(Len(svSQLBody)), "0", 4) & svSQLBody & "________"
      
      aStrucStack(iStrucLevel) = Array(10, Len(Res) + 1, "", Str(Len(Res) - 8)) 'цикл
      iStrucLevel = iStrucLevel + 1
      
     'Если за полем идет \par то отбрасываем его
      nState = 3
      
      
     Case "skip"
      skipConst = True
     Case "endskip"
       'Ни чего не делает просто ограничение окончания скипа, может забрать с собой перевод строки
       If LCase(Trim(sOpt)) = "skip_lf" Then nState = 3
      
     Case "endscan"
     
      iStrucLevel = iStrucLevel - 1
      If aStrucStack(iStrucLevel)(0) <> 10 Then Err.Raise 8001, , "Endscan нарушает структуру программы"
      
      'разносим адреса
      If aStrucStack(iStrucLevel)(2) <> "" Then
       For Each adr In Split(aStrucStack(iStrucLevel)(2), ",")
        Res = InsertAddress(Res, Len(Res) + 1, CInt(adr))
       Next
      End If
            
      Res = Res & "NEXT" & LPad(Hex(aStrucStack(iStrucLevel)(1)), "0", 8)
         
      If aStrucStack(iStrucLevel)(3) <> "" Then
       For Each adr In Split(aStrucStack(iStrucLevel)(3), ",")
        Res = InsertAddress(Res, Len(Res) + 1, CLng(adr))
       Next
      End If
         
     Case "if"
      
      
      aStrucStack(iStrucLevel) = Array(0, Str(Len(Res) + 1), "") 'if then or elif
      iStrucLevel = iStrucLevel + 1
      Res = Res & "JMPF" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt & "________"

      
     'Если за полем идет \par то отбрасываем его
      nState = 3

      
     Case "elif"
      If Not aStrucStack(iStrucLevel - 1)(0) = 0 Then Err.Raise 8003, , "Elif нарушает структуру программы"
     
     
      Res = Res & "GOTO________"
      If aStrucStack(iStrucLevel - 1)(2) <> "" Then aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) + ","
      aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) & Str(Len(Res) + 1)
      
      Res = InsertAddress(Res, Len(Res) + 1, CLng(aStrucStack(iStrucLevel - 1)(1)))
      
      Res = Res & "JMPF" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt & "________"
      aStrucStack(iStrucLevel - 1)(1) = Str(Len(Res) + 1)

     'Если за полем идет \par то отбрасываем его
      nState = 3

      
     Case "else"
      If Not aStrucStack(iStrucLevel - 1)(0) = 0 Then Err.Raise 8003, , "Else нарушает структуру программы"
      aStrucStack(iStrucLevel - 1)(0) = 1
      
      Res = Res & "GOTO"
      If aStrucStack(iStrucLevel - 1)(2) <> "" Then aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) + ","
      aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) & Str(Len(Res) + 1)
      Res = Res & "________"
      
      Res = InsertAddress(Res, Len(Res) + 1, CLng(aStrucStack(iStrucLevel - 1)(1)) - 1)
      aStrucStack(iStrucLevel - 1)(1) = ""
      
     'Если за полем идет \par то отбрасываем его
      nState = 3
      
     Case "endif"
      If Not InSet(aStrucStack(iStrucLevel - 1)(0), 0, 1) Then Err.Raise 8003, , "Endif нарушает структуру программы"
      
      iStrucLevel = iStrucLevel - 1
      If aStrucStack(iStrucLevel)(1) <> "" Then
       If aStrucStack(iStrucLevel)(2) <> "" Then aStrucStack(iStrucLevel)(2) = aStrucStack(iStrucLevel)(1) & ","
       aStrucStack(iStrucLevel)(2) = aStrucStack(iStrucLevel)(2) & aStrucStack(iStrucLevel)(1)
      End If
      
      For Each adr In Split(aStrucStack(iStrucLevel)(2), ",")
       Res = InsertAddress(Res, Len(Res) + 1, CLng(adr) - 1)
      Next

      'Обрабатываем опциональный параметр пропуска перевода строки
      If LCase(Trim(sOpt)) = "skip_lf" Then nState = 3
      
         
     Case "calc"
       Res = Res & "CALC" & LPad(Hex(Len(sTXT)), "0", 3) & sTXT
     Case "next"
       Res = Res & "CONT" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
     Case "fld", "f"
      If SFMT = "\no" Then
       Res = Res & "PRVL" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
      Else
       skipConst = False
       SFMT = "{" & Trim(SFMT) '& "  "
       Res = Res & "PRNT" & LPad(Hex(Len(SFMT)), "0", 8) & SFMT
       Res = Res & "PRVL" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
       Res = Res & "PRNT00000001}"
      End If
      
      
      
     Case "setm"
      sOpt = "calc(" & sOpt & ";" & Len(Res) + 1 & ")"
      MetList = MetList & "CALC" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
     Case "jump"
      Res = Res & "JUMP" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
     Case "call"
      Res = Res & "CALL" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
     Case "ret"
      Res = Res & "RETC"
      
      
     Case Else
      Err.Raise 1000, "", "Не известная комманда: " & sFnc
     End Select
   End If
   CP = tp
  Else
   tp = Len(ts) + 1
   Res = Res & "PRNT" & LPad(Hex(tp - CP), "0", 8) & Mid(ts, CP, tp - CP)
   CP = Len(ts) + 1
  End If
  
  
  
 Case 3
    Dim svToken, nvLength
    tp = CP
    
    svToken = RTF_SkipPar(ts, tp)
    
    If Trim(svToken) = "c\par" Then
    
      If Mid(ts, tp - 5, 1) = " " Then
       nvLength = 6
      ElseIf InSet(Mid(ts, tp - 5, 1), vbCr, vbLf) And Mid(ts, tp, 1) = " " Then
       tp = tp + 1
       nvLength = 6
      Else
       nvLength = 5
      End If
      
      ts = Mid(ts, 1, tp - nvLength) & Mid(ts, tp)

      nState = 0
    Else
      nState = 1
    End If

 End Select
  
 If nState = 2 Then
   tp = CP
   svToken = RTF_SkipPar(ts, tp)
   nState = 1
 End If
  
 If nState = 1 Then
   If svToken = "c\field" Then
     CP = InStr(CP, ts, "{\field", vbTextCompare)
   End If
   nState = 0
 End If

  
 Loop
 
 Res = Res & "ENDT"
 sOpt = Len(Res)
 sOpt = "CALL00d" & LPad(Trim(sOpt & ""), "0", 13)
  
 If MetList <> "" Then Res = sOpt & Mid(Res, Len(sOpt) + 1) & MetList & "RETC"
 
 If iStrucLevel > 1 Then
  MsgBox IIf(aStrucStack(iStrucLevel - 1)(0) <> 10, "В ходе разбора файла обнаружен не закрытый блок IF. ", "") & IIf(aStrucStack(iStrucLevel - 1)(0) = 10, "В ходе разбора файла обнаружен не закрытый блок SCAN.", "")
  Exit Function
 End If

 PrepareRTF = Res
' Set tf = fso.CreateTextFile(sFile & "_", True)
' tf.Write Res
' tf.Close
 
End Function

Private Function ToBool(Value As Variant, Optional bRaise As Boolean = True) As Variant
'Перобразнует значение в булевый формат и возвращет его значение
'#param Value: Значение
'#param bRaise: Что делать если значение не булево
' {*} True - Генерировать ошибку
' {*} False - Вернуть Empty

 
 On Error GoTo NoBool
 ToBool = CBool(Value)
 Exit Function
NoBool:
 ToBool = Empty
 If bRaise Then
   On Error GoTo 0
   Err.Raise 1000, , "ToBool: значение {" & Value & "} не логического типа"
 End If
End Function

Private Function DumpContext(ByRef ParamList As Variant) As String
'Возвращает содержимое контекста в виде листинга
'#param ParamList: Словарь контекста

  Dim key
  DumpContext = "КОНТЕКСТ:"
  For Each key In ParamList.Keys
   If Not InSet(LCase(key), "%", "now", "date") Then
     If Not IsArray(ParamList(key)) Then
       DumpContext = DumpContext & vbCrLf & key & " = " & Mid(ParamList(key), 1, 128) & IIf(Len(ParamList(key)) > 128, "...", "")
     End If
   End If
  Next

End Function


Public Function GetValue(Formula As Variant, ByRef ParamList As Variant, ByRef StartPos As Long) As Variant
'Внутренняя функция для формирования отчета. Разбирает варыжение и возвращает его значение. Строки задаются в двойных ковычках.
'Значения переменных должны находиться в текущем контексте или находиться при помощи функции eval. Стартовые пробелы пропускаются.
'Для получения более подробной информации см. справку.
'#param Formula: Текст выражения
'#param ParamList: Текущий контекст
'#param StartPos: Текущее смещение в выражении

Dim StopSym As String
Dim CP As Integer
Dim sFnc As String
Dim aArg(16) As Variant, iArg As Integer
Dim nvSP As Integer, Item
      Dim objXML, objDocElem
      Dim byteStorage() As Byte
      Dim BCWidth

'Пропускаем не значищие пробелы
Do While Mid(Formula, StartPos, 1) = " "
 StartPos = StartPos + 1
Loop

nvSP = StartPos

'список терминаторов
StopSym = "(;)"

On Error GoTo OnError

If Mid(Formula, StartPos, 1) = """" Then 'строковая константа
 StartPos = StartPos + 1
 Do While StartPos <= Len(Formula)
  If Mid(Formula, StartPos, 2) = """""" Then
   sFnc = sFnc & """"
   StartPos = StartPos + 2
  ElseIf Mid(Formula, StartPos, 1) = """" Then
   GetValue = sFnc
   StartPos = StartPos + 1
   Exit Do
  Else
   sFnc = sFnc + Mid(Formula, StartPos, 1)
   StartPos = StartPos + 1
  End If
 Loop
 'GetValue = Replace(Mid(GetValue, 2, Len(GetValue) - 1), """""", """")
 
ElseIf Mid(Formula, StartPos, 1) = ")" Or StartPos > Len(Formula) Then
 GetValue = Null
 Exit Function
Else
 sFnc = ""
 Do While InStr(1, StopSym, Mid(Formula, StartPos, 1)) = 0 And StartPos <= Len(Formula)
  sFnc = sFnc + Mid(Formula, StartPos, 1)
  StartPos = StartPos + 1
 Loop
 
 sFnc = Trim(sFnc)

 
 If sFnc = "" Then Err.Raise 1000, , "Ожидается какое либо значение."
 
 If Mid(Formula, StartPos, 1) = "(" Then 'Функция
  iArg = 0
  StartPos = StartPos + 1 'пропускаем (
  Do While True 'Собираем аргументы
   aArg(iArg) = GetValue(Formula, ParamList, StartPos)
   If Mid(Formula, StartPos, 1) = ")" Then
    StartPos = StartPos + 1
    Exit Do
   End If
   If IsNull(aArg(iArg)) Then Exit Do
   iArg = iArg + 1 ' пропускаем точку с запятой
   StartPos = StartPos + 1
  Loop
  
  Select Case LCase(sFnc)
  Case "+", "plus"
   GetValue = 0
   For Each Item In aArg
    If IsNull(Item) Or IsEmpty(Item) Then Exit For
    GetValue = GetValue + Item
   Next
  Case "-", "minus"
   GetValue = aArg(0) - aArg(1)
  Case "*", "mul"
   GetValue = 1
   For Each Item In aArg
    If IsNull(Item) Or IsEmpty(Item) Then Exit For
    GetValue = GetValue * Item
   Next
  Case "/", "div"
   GetValue = aArg(0) / aArg(1)
  Case "\", "idiv"
   GetValue = aArg(0) \ aArg(1)
  Case "mod"
   GetValue = aArg(0) Mod aArg(1)
  Case "&", "concat"
   For Each Item In aArg
    If IsNull(Item) Or IsEmpty(Item) Then Exit For
    GetValue = GetValue & Item
   Next
  Case "=", "eq"
   GetValue = (aArg(0) = aArg(1))
  Case "<", "ls"
   GetValue = (aArg(0) < aArg(1))
  Case ">", "gr"
   GetValue = (aArg(0) > aArg(1))
  Case "<=", "le"
   GetValue = (aArg(0) <= aArg(1))
  Case ">=", "ge"
   GetValue = (aArg(0) >= aArg(1))
  Case "<>", "!=", "ne"
   GetValue = (aArg(0) <> aArg(1))
  Case "or"
   GetValue = False
   For Each Item In aArg
    If IsNull(Item) Or IsEmpty(Item) Then Exit For
    If ToBool(Item) Then
      GetValue = True
      Exit For
    End If
   Next
  Case "iif"
   GetValue = IIf(aArg(0) <> 0, aArg(1), aArg(2))
  Case "and"
   GetValue = True
   For Each Item In aArg
    If IsNull(Item) Or IsEmpty(Item) Then
      GetValue = False
      Exit For
    End If
    If Not ToBool(Item) Then
      GetValue = False
      Exit For
    End If
   Next
  Case "xor"
   GetValue = False
   For Each Item In aArg
    If IsNull(Item) Or IsEmpty(Item) Then Exit For
    GetValue = GetValue Xor ToBool(Item)
   Next
  Case "not"
   GetValue = Not ToBool(Item)
  Case "isnull"
   GetValue = IsNull(aArg(0))
  Case "open"
    Dim objStream, fso
    Set fso = CreateObject("scripting.FileSystemObject")
    
    If Left(aArg(0), 2) = ".\" Then
      aArg(0) = GetPath(CurrentDb.Name) & Mid(aArg(0), 3)
    End If
    
    If Not fso.FileExists(aArg(0)) Then
      GetValue = "Файл '" & aArg(0) & "' не найден."
    Else
      
      Set objStream = CreateObject("ADODB.Stream")
      If aArg(1) & "" <> "" Then
       objStream.Charset = aArg(1)
       objStream.Type = 2 'adTypeText
      Else
       objStream.Type = 1 ' adTypeBinary
      End If
      
      objStream.Open
      
      objStream.LoadFromFile (aArg(0))
      GetValue = objStream.Read()
      Set objStream = Nothing
    End If
    Set fso = Nothing
  Case "attach" ' 1 параметр - это поле attachment, второе поле - это маска поиска
    Dim rsFiles, objRegExp, FileName
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = False
    objRegExp.IgnoreCase = True
    objRegExp.Multiline = False
    'Фильтр по умолчанию настроен на картинки
    If aArg(1) <> "" Then objRegExp.Pattern = aArg(1) Else objRegExp.Pattern = ".+\.(jpg|jpeg|png|emf)$"
    
    For Each FileName In aArg(0)(0).Keys
      If objRegExp.test(FileName) Then
          Set objXML = CreateObject("MSXml2.DOMDocument")
          Set objDocElem = objXML.createElement("Base64Data")
          objDocElem.DataType = "bin.hex"
          objDocElem.nodeTypedValue = aArg(0)(0)(FileName)
          objDocElem.Text = Mid(objDocElem.Text, 41)
          GetValue = objDocElem.nodeTypedValue
          Set objDocElem = Nothing
          Set objXML = Nothing
          Exit For
      End If
    Next
  Case "rtfimg"
    If IsNull(aArg(0)) Then
      GetValue = ""
    ElseIf TypeName(aArg(0)) <> "byte()" Then
      GetValue = "{Здесь должно быть изображение: " & aArg(0) & "}"
    Else
      byteStorage = aArg(0)
      GetValue = PictureDataToRTF(byteStorage, aArg(1), aArg(2))
      
    End If
    
   
  Case "rel"
   sFnc = Trim(aArg(0))
   If ParamList.Exists(sFnc) Then
    GetValue = ParamList(sFnc)
    Exit Function
   Else
    Err.rise 1001
   End If
   
   If ParamList.Exists(GetValue) Then
    GetValue = ParamList(GetValue)
    Exit Function
   Else
    Err.rise 1001
   End If
 
 'дальше описываем остальные функции
  Case "calc"
   ParamList("" & aArg(0)) = aArg(1)
   GetValue = "" & aArg(0)
  
  Case Else
   On Error GoTo OnNoFunction
   GetValue = Application.Run(sFnc, ParamList, aArg)
  End Select
  
 ElseIf IsNumeric(sFnc) Then 'Если похоже на число то возварщаем число
  GetValue = CDbl(sFnc)
  Exit Function
 Else
  If ParamList.Exists(sFnc) Then 'Иначе считаем переменной и ищем в списке
   GetValue = ParamList(sFnc)
   Exit Function
  Else
   On Error GoTo ParamNotFound
   GetValue = Application.Eval(sFnc)
   ParamList(sFnc) = GetValue
   Exit Function
ParamNotFound:
   Err.Raise 1001, , "Параметр '" & sFnc & "' не найден в списке"
  End If
 End If

End If
Exit Function

OnError:
  Dim sErrorMsg As String
  sErrorMsg = "Ошибка в формуле {" & Formula & "} в позиции " & nvSP & "." & vbCrLf & Err.Description
  On Error GoTo 0
  Err.Clear
  sErrorMsg = sErrorMsg & vbCrLf & vbCrLf & DumpContext(ParamList)
  Err.Raise 1001, , sErrorMsg
OnNoFunction:
  sErrorMsg = "Ошибка в формуле {" & Formula & "} в позиции " & nvSP & "." & vbCrLf & "Не известная функция [" & sFnc & "]"
  On Error GoTo 0
  Err.Clear
  sErrorMsg = sErrorMsg & vbCrLf & vbCrLf & DumpContext(ParamList)
  Err.Raise 1001, , sErrorMsg
End Function


Public Function GetTemplate(idReport As Long) As Variant()
'Внутренняя функция для формирования отчета. Извлекает из хранилища шаблон по его коду.
'Так же производится проверка, если исходный файл существует и его дата изменения больше чем у сохраненного шаблона, то шаблон будет обновлен.
'#param idReport: Код шаблона

 Dim fso
 Dim objF
 Dim tRep As Recordset
 Dim sPath As String
 Dim sPathOrig As String, sExtension As String
 
 Set fso = CreateObject("scripting.FileSystemObject")
 
 Set tRep = CurrentDb.OpenRecordset("t_Rep", dbOpenDynaset)
 tRep.FindFirst "id = " & idReport
  
 If tRep.NoMatch Then Err.Raise 1000, , "Не найден шаблон с кодом " & idReport
 
 sPathOrig = tRep("sOrignTemplate")
 If Left(sPathOrig, 2) = ".\" Then sPathOrig = CurrentProject.Path & Mid(sPathOrig, 2)
 sExtension = LCase(fso.GetExtensionName(sPathOrig))
  
 If fso.FileExists(sPathOrig) Then
   Set objF = fso.GetFile(sPathOrig)
  
   If Nz(tRep("dEditTemplate"), Now) <> objF.DateLastModified Then
     tRep.Edit
     Select Case LCase(sExtension)
      Case "rtf"
       tRep("clTemplate") = PrepareRTF(sPathOrig)
     End Select
          
     tRep("dEditTemplate") = objF.DateLastModified
     tRep.Update
   End If
   Set objF = Nothing
   End If
 GetTemplate = Array(sExtension, tRep("clTemplate") & "")
 tRep.Close
 Set tRep = Nothing
 
End Function

Function fncConvertTxtToRTF(Text As String) As String
'Внутренняя функция для формирования отчета. Конвертирует текст в корректный RTF блок. Если начинается с `{\*\shppict`, то оставляет как есть
'#param Text: Текст

 Dim i As Long, ch As String
 If LCase(Left(Text, 11)) = "{\*\shppict" Then
    fncConvertTxtToRTF = Text 'Не меняем
 Else
    fncConvertTxtToRTF = " "
    For i = 1 To Len(Text)
     ch = Mid(Text, i, 1)
     Select Case Asc(ch)
     Case Asc("{"), Asc("}"), Asc("\")
      fncConvertTxtToRTF = fncConvertTxtToRTF & "\" & ch
     Case &H0 To &HF
      fncConvertTxtToRTF = fncConvertTxtToRTF & "\'0" & Hex(Asc(ch))
     Case &H10 To &H1F, &H80 To &HFF
      fncConvertTxtToRTF = fncConvertTxtToRTF & "\'" & Hex(Asc(ch))
     Case Else
      fncConvertTxtToRTF = fncConvertTxtToRTF & ch
     End Select
    Next
 End If
End Function


Public Function PictureDataToRTF(PictureData, nWidth, nHeight)
'Внутренняя функция для формирования отчета. Конвертирует изображение в корректный RTF блок.
'#param PictureData: Байтовый массив с изображением
'#param nWidth: Целевая ширина картинки
'#param nHeight: Целевая высота картинки


  PictureDataToRTF = "{\*\shppict{\pict\picwgoal" & Int(nWidth * 56.6929133858) & "\pichgoal" & Int(nHeight * 56.6929133858)
  
  Select Case GetTypeContent(PictureData)
   Case "jpg": PictureDataToRTF = PictureDataToRTF & "\jpegblip" & vbCrLf
   Case "png": PictureDataToRTF = PictureDataToRTF & "\pngblip" & vbCrLf
   Case "emf": PictureDataToRTF = PictureDataToRTF & "\emfblip" & vbCrLf
   Case "wmf": PictureDataToRTF = PictureDataToRTF & "\wmetafile7" & vbCrLf
  End Select
  
  
  Dim objDocElem
  Set objDocElem = CreateObject("MSXml2.DOMDocument").createElement("Base64Data")
  objDocElem.DataType = "bin.hex"
  objDocElem.nodeTypedValue = PictureData
  PictureDataToRTF = PictureDataToRTF & objDocElem.Text & "}}"
  Set objDocElem = Nothing

End Function

Public Function MakeReport(ts As String, ByRef OutStream As Variant, ByRef p_Dic As Variant) As String
'Внутренняя функция для формирования отчета. Непосредственное формирование отчета по скомпилированному шаблону
'#param ts: Шаблон
'#param OutStream: Выходной поток куда будет записан сформированный документ
'#param p_Dic: Контекст

 Dim fso
 Dim tFile 'As TextStream
 'Dim ts As String
 Dim PC As Long, iCnt As Long, iCnt2 As Long
 Dim dic
 Dim sSQL As String
 
 Dim aRecordSet As Variant
 
 Dim aRetCall(128) As Long
 Dim iRC As Integer
 Dim sfncConvert As String
 Dim sValue As Variant, sName, sAlias As String, fld
 
 Set fso = CreateObject("scripting.FileSystemObject")
 
 If p_Dic Is Nothing Then
  Set dic = CreateObject("Scripting.Dictionary")
 Else
  Set dic = p_Dic
 End If
 'dic.CompareMode = 1
 
 iRC = 0
 
 If dic("extension") <> "" Then
  sfncConvert = "fncConvertTxtTo" & dic("extension")
 End If
 
 On Error Resume Next
 sValue = Application.Run(sfncConvert, "")
 If Err Then sfncConvert = ""
 On Error GoTo 0
  
 PC = 1
 
'<DOC>
'# Описание внутреннего формата
'
'Все блоки начинаются с строки длиной в 4 символа, которые определяют тип блока. Ниже описаны все используемые блоки:
'
 
 
 Do While PC <= Len(ts)
  Select Case UCase(Mid(ts, PC, 4))
   Case "PRNT"
   
'<DOC>
'`PRNT N[8] S[N]`
'Добавляет текстовый литерал
'- `N` - длина текстового блока, записанная в виде 8 чисел в шестнадцетиричном виде
'- `S` - Строка длиной заданной в N

    iCnt = CLng("&h" & Mid(ts, PC + 4, 8))
    If Not OutStream Is Nothing Then OutStream.Write Mid(ts, PC + 12, iCnt) Else MakeReport = MakeReport & Mid(ts, PC + 12, iCnt)
    PC = PC + iCnt + 12
   Case "PRVL"
   
'<DOC>
'`PRVL N[3] V[N]`
'Добавляет в документ результат выражения
'- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
'- `V` - Выражение длиной заданной в N

    iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
    
    sValue = Mid(ts, PC + 7, iCnt)
    
    On Error Resume Next
    sValue = Application.Eval(sValue)
    If Err Then
      On Error GoTo 0
      Dim StartPos As Long, sFormula As Variant, sFormat As String
      StartPos = 1
      sFormula = sValue
      sValue = GetValue(sFormula, dic, StartPos)
      StartPos = StartPos + 1
      If StartPos < Len(sFormula) Then
        sFormat = GetValue(sFormula, dic, StartPos)
        sValue = Format(sValue, sFormat)
      End If
    Else
      On Error GoTo 0
    End If
   
    sValue = Nz(sValue, "")
    If sfncConvert <> "" Then sValue = Application.Run(sfncConvert, sValue)
    If Not OutStream Is Nothing Then OutStream.Write sValue Else MakeReport = MakeReport & sValue
    PC = PC + iCnt + 7
   Case "CALC"
'<DOC>
'`CALC N[3] V[N]`
'Выполняет выражение, но значение игнорируется
'- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
'- `V` - Выражение длиной заданной в N

    iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
    GetValue Mid(ts, PC + 7, iCnt), dic, 1
    PC = PC + iCnt + 7
   Case "GOTO"
    PC = CLng("&h" & Mid(ts, PC + 4, 8))
    
   Case "JUMP"
'<DOC>
'`JUMP N[3] A[N]`
'Выполняет выражение, значение используется как новый адрес для выполнения
'- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
'- `A` - Выражение длиной заданной в N, в котором хранится новый адрес
    iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
    PC = GetValue(Mid(ts, PC + 7, iCnt), dic, 1)
    If PC = 0 Then Err.Raise 1111
    
   Case "CALL"
'<DOC>
'`CALL N[3] A[N]`
'Выполняет выражение, значение используется как новый адрес для выполнения, при этом в стек вызовов добавляется адрес за текущей конструкцией
'- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
'- `A` - Выражение длиной заданной в N, в котором хранится новый адрес
    iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
    aRetCall(iRC) = PC + 8 + iCnt
    iRC = iRC + 1
    PC = GetValue(Mid(ts, PC + 7, iCnt), dic, 1)
    If PC = 0 Then Err.Raise 1111
   Case "RETC"
'<DOC>
'`RETC`
'Извлекает из стека сохраненный адрес и передает управленение на него
    If iRC = 0 Then Err.Raise 1110
    iRC = iRC - 1
    PC = aRetCall(iRC)
    
   Case "JMPF"
'<DOC>
'JMPF N[3] V[N] J[8]
'Условный прыжок выполняет выражение, если значение равно Истина, то управление передается на инструкцию за текущей.
'В противном случае осуществляет прыжок по адресу указанному в поле `J`
'- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
'- `V` - Выражение длиной заданной в N
'- `J` - Новый адрес если выражение равно Ложь

    iCnt = CLng("&h" & Mid(ts, PC + 4, 3))
    If Not ToBool(GetValue(Mid(ts, PC + 7, iCnt), dic, 1)) Then
     PC = CLng("&h" & Mid(ts, PC + 7 + iCnt, 8))
    Else
     PC = PC + 15 + iCnt
    End If
'<DOC>
'`ENDT`
'Метка конца шаблона
   Case "ENDT"
    Exit Do
    
   Case "OPRS"
'<DOC>
'`OPRS I[3] N[I] J[4] S[J] E[8]`
'Открывает набор данных.
'- `I` - длина выражения для получения имени набора данных, записанная в виде 3 чисел в шестнадцетиричном виде
'- `N` - Выражение длиной заданной в I, результат которого используется как имя набора данных
'- `J` - длина выражения для получения источника данных (SQL), записанная в виде 4 чисел в шестнадцетиричном виде
'- `S` - Выражение длиной заданной в J, результат которого используется как источник данных
'- `E` - Новый адрес если набор данных не содержит данных
    
    iCnt = CLng("&h" & Mid(ts, PC + 4, 3))
    sName = Trim(GetValue(Mid(ts, PC + 7, iCnt), dic, 1))
    
    iCnt2 = CLng("&h" & Mid(ts, PC + iCnt + 7, 4))
    sSQL = Trim(GetValue(Mid(ts, PC + iCnt + 11, iCnt2), dic, 1)) 'получаем текст из шаблона
    sSQL = FilterFmt(sSQL, dic) 'подставляем переменные
    
    'Новый набор данных
    aRecordSet = Array(sName, Empty, dic("@SYS_CurrentRecordSet"))
    
    On Error Resume Next
    
'ToDo: Добавить возможность брать готовый наборы данных из запросов или форм по префиксу @

    Set aRecordSet(1) = CurrentDb.OpenRecordset(sSQL)
    
    If Err Then 'Формируем текст ошибки
      On Error GoTo 0
      Dim sErrorMsg As String
      sErrorMsg = "Ошибка при открытии курсора {" & sName & "}." & vbCrLf & Err.Description
      Err.Clear
      sErrorMsg = sErrorMsg & vbCrLf & vbCrLf & sSQL & vbCrLf & vbCrLf & DumpContext(p_Dic)
      MsgBox sErrorMsg
      Err.Raise 1001, , sErrorMsg
    End If
    On Error GoTo 0
    
    
    If aRecordSet(1).EOF Then
     aRecordSet(1).Close
     Set aRecordSet(1) = Nothing
     PC = CLng("&h" & Mid(ts, PC + 11 + iCnt + iCnt2, 8))
     aRecordSet = aRecordSet(2)
    Else
     dic("@SYS_CurrentRecordSet") = aRecordSet 'Если записей нет то ни чего не меняем
     dic(sName & ".rownum") = 0
     PC = PC + iCnt + 19 + iCnt2
     FetchRow dic
    End If
    
   Case "NEXT"
'<DOC>
'`NEXT A[8]`
'Считывает очередную строчку из набора данных и передает управление на начало цикла.
'Если данные кончились, то передает управление на конструкцию после текущей.
'- `A` - Адрес начала цикла

    If FetchRow(dic) Then
      PC = CLng("&h" & Mid(ts, PC + 4, 8))
    Else
      aRecordSet(1).Close
      Set aRecordSet(1) = Nothing
      aRecordSet = aRecordSet(2)
      dic("@SYS_CurrentRecordSet") = aRecordSet
      PC = PC + 12
    End If
    

   Case "CONT"
'<DOC>
'`CONT N[3] V[N]`
'Переход к следующей записи внутри цикла
'- `N` - Длина выражения
'- `V` - Выражение, результатом которого должно быть имя набора данных. Если результат - пустая строка, то используется текущий набор данных

'ToDo: Для единообразия перенести в CALC и использовать как функцию

    iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
    sValue = Mid(ts, PC + 7, iCnt)
    PC = PC + iCnt + 7
    On Error Resume Next
    sValue = Application.Eval(sValue)
    If Err Then
      On Error GoTo 0
      sValue = GetValue(sValue, dic, 1)
    Else
      On Error GoTo 0
    End If
    
    If Not FetchRow(dic, sValue) Then
      'Очиcтим все переменные с указанным префиксом
      sValue = UCase(sValue & ".")
      Dim l, sKey
      l = Len(sValue)
      For Each sKey In dic.Keys()
        If UCase(Mid(sKey, 1, l)) = sValue Then
          If UCase(sKey) <> sValue & "EOF" Then dic(sKey) = Empty
        End If
      Next
      
    End If
   Case Else
    Err.Raise 1001, , "Шаблон поломался :("
  End Select
 Loop

End Function

Public Function FetchRow(ByRef pDic, Optional ByVal pCursorName As String = "")
'Внутренняя функция для формирования отчета. Извлекает из курсора очередную строку и обновляет значения в контексте.
'#param pDic: Текущий контекст
'#param pCursorName: Имя курсора. Если не задан то считыается текущий курсор


  Dim vRecordSet, vCursorName, vFiles, tmpdic, fld
  
  vRecordSet = pDic("@SYS_CurrentRecordSet")
  If pCursorName = "" Then
    vCursorName = vRecordSet(0)
    Set vRecordSet = vRecordSet(1)
  Else
    vCursorName = pCursorName
    Do While True
      If UCase(vRecordSet(0)) = UCase(pCursorName) Then
        Set vRecordSet = vRecordSet(1)
        Exit Do
      End If
      If Not IsArray(vRecordSet(2)) Then
        'Не нашли курсор с указанным именем
        FetchRow = False
        Exit Function
      End If
      vRecordSet = vRecordSet(2)
    Loop
  End If
  If Not vRecordSet.EOF Then
    For Each fld In vRecordSet.Fields
      If IsObject(fld.Value) Then
        If TypeName(fld.Value) = "Recordset2" Then
         
          Set tmpdic = CreateObject("Scripting.Dictionary")
          tmpdic.CompareMode = 1
          
          Set vFiles = fld.Value
          While Not vFiles.EOF
            tmpdic(vFiles.Fields("FileName").Value) = vFiles.Fields("FileData").Value
            vFiles.MoveNext
          Wend
          vFiles.Close
          Set vFiles = Nothing
          
          pDic(vCursorName & "." & fld.Name) = Array(tmpdic)
        End If
      Else
        pDic(vCursorName & "." & fld.Name) = fld.Value
      End If
    Next
    vRecordSet.MoveNext
    pDic(vCursorName & ".EOF") = vRecordSet.EOF
    pDic(vCursorName & ".rownum") = pDic(vCursorName & ".rownum") + 1
    Set vRecordSet = Nothing
    FetchRow = True
  Else
    FetchRow = False
    pDic(vCursorName & ".EOF") = True
  End If
  Set vRecordSet = Nothing
End Function

Public Function GetPath(FullPath As String) As String
'Возвращает имя директории файла
'#param FullPath: Полное имя
  Dim lngCurrPos, lngLastPos As Long
  Do
    lngLastPos = lngCurrPos
    lngCurrPos = InStr(lngLastPos + 1, FullPath, "\")
  Loop Until lngCurrPos = 0
  If lngLastPos <> 0 Then GetPath = Left(FullPath, lngLastPos)
End Function

Public Function GetFile(FullPath As String) As String
'Возвращает имя файла
'#param FullPath: Полное имя
  Dim lngCurrPos, lngLastPos As Long
  Do
    lngLastPos = lngCurrPos
    lngCurrPos = InStr(lngLastPos + 1, FullPath, "\")
  Loop Until lngCurrPos = 0
  If lngLastPos <> 0 Then GetFile = Right$(FullPath, Len(FullPath) - lngLastPos)
End Function

Public Function GetExt(FullPath As String) As String
'Возвращает расширение файла
'#param FullPath: Полное имя
  Dim lngCurrPos, lngLastPos As Long
  Do
    lngLastPos = lngCurrPos
    lngCurrPos = InStr(lngLastPos + 1, FullPath, ".")
  Loop Until lngCurrPos = 0
  If lngLastPos <> 0 Then GetExt = Right$(FullPath, Len(FullPath) - lngLastPos + 1)
End Function

Function GetTypeContent(ByRef tpData)
'Определение формата изображения по его внутренней структуре. Возвращает следующие значения:
' {*} jpg
' {*} png
' {*} emf
' {*} wmf
' {*} не распознан - если не удалось определить формат изображения
'#param p_Dic: Массив байт картинки

  If IsNull(tpData) Then
    GetTypeContent = ""
  ElseIf VarType(tpData) = 8209 Then
    If UBound(tpData) > 4 Then
      If tpData(0) = &HFF Then
        If tpData(1) = &HD8 And tpData(2) = &HFF And tpData(3) = &HE0 Then
          GetTypeContent = "jpg"
        Else
          GetTypeContent = "не распознан"
        End If
      ElseIf tpData(1) = &H50 Then
        If tpData(2) = &H4E And tpData(3) = &H47 Then
          GetTypeContent = "png"
        Else
          GetTypeContent = "не распознан"
        End If
      ElseIf tpData(0) = &H1 Then
        If tpData(1) = &H0 And tpData(2) = &H0 And tpData(3) = &H0 And tpData(&H28) = &H20 And tpData(&H29) = &H45 And tpData(&H2A) = &H4D And tpData(&H2B) = &H46 Then
          GetTypeContent = "emf"
        ElseIf tpData(1) = &H0 And (tpData(5) = &H1 Or tpData(5) = &H3) And tpData(4) = &H0 Then
          GetTypeContent = "wmf"
        Else
          GetTypeContent = "не распознан"
        End If
      End If
    Else
      GetTypeContent = "не распознан"
    End If
  Else
    GetTypeContent = "не распознан"
  End If

End Function
      

Function FilterFmt(Text As String, ByRef p_Dic As Variant) As String
'Производит замену в тексте подстановочных сиволов на значения заданные в словаре.
'Подстановочные символы обрамляются символом `%`. Если нужно вывести символ как есть то его необходимо удвоить `%%`. Значение ключа ищется в словаре.
'Если такого значения нет, то будет предпринята попытка получить значение через Eval
'#param Text: Исходный тест с подстановками.
'К подстановочному значению можно применить операции форматирования, для этого нужно после ключа через символ `;` указать способ формтирования.
'Доступные варианты форматирования:
' {*} stdf:<имя_фильтра> - операция применения фильтра. см. отдельную справку по форматированию фильтров.
' {*} sqldate:<значение_по_умолчанию> - Форматирует дату как SQL литерал. Если значение является Null, то берется <значение_по_умолчанию> как есть. Значение по умолчанию не обязательно и опускается вместе с символом `:`
' Специальный подстановочный имена полей
' {*} fnc<Имя_функции>:<Ключ в словаре> - Применение пользовательской функции для форматирования значения
' {*} get:<Выражение> - <Выражение> пропускается через функцию GetValue и подставляется значение
' Если ни один вариант выше не подошел, то используется как выражение формата для функции Format
'#param p_Dic: Словарь с значениями подстановок

 'Dim FilterFmt As String
 Dim p As Long, pn As Long, smid As String, pt As Long
 Dim SFMT, sKey As String
 Dim idOperation As String, sPrefix As String
 Dim sOperand, sOperand2
 
 FilterFmt = Text
 p = 1
 
 If Not p_Dic.Exists("%") Then p_Dic.Add "%", "%"
 
 
 Do While True
  p = InStr(p, FilterFmt, "%")
  If p = 0 Then Exit Do
  If Mid(FilterFmt, p + 1, 1) <> "%" Then Exit Do
  FilterFmt = Left(FilterFmt, p) & Mid(FilterFmt, p + 2)
  p = p + 1
 Loop
 
 'p = InStr(1, FilterFmt, "%")
 
 On Error GoTo ErrorMet
 Do While p > 0
  pn = InStr(p + 1, FilterFmt, "%")
  
  Do While True
   pn = InStr(p + 1, FilterFmt, "%")
   If pn = 0 Then Exit Do
   If Mid(FilterFmt, pn + 1, 1) <> "%" Then Exit Do
   FilterFmt = Left(FilterFmt, p) & Mid(FilterFmt, p + 2)
   p = pn
  Loop
  
  If pn < 1 Then
   Exit Function
  ElseIf pn - p = 1 Then
   FilterFmt = Left(FilterFmt, p) & Mid(FilterFmt, pn + 1)
   p = p + 1
  Else
   smid = Mid(FilterFmt, p + 1, pn - p - 1)
   pt = InStr(1, smid, ";")
   If pt > 0 Then 'применяется операция форматирования вывода
    
    SFMT = Mid(smid, pt + 1)
    sKey = Trim(Mid(smid, 1, pt - 1))
    
    If LCase(Left(SFMT, 5)) = "stdf:" Then 'Формирование фильтра по полю
     SFMT = Mid(SFMT, 6)

     If Not p_Dic.Exists(sKey & ".oper") Or p_Dic(sKey & ".oper") = 0 Then 'Если операция не указана то ни чего не выводим
      smid = ""
     Else
        idOperation = p_Dic(sKey & ".oper")
        sPrefix = p_Dic(sKey & ".type") 'Базовый тип операнда
        
        If idOperation = opIN Or idOperation = opNIN Then 'In, Not in
        
         Dim sList As String
         sList = ""
         Dim i As Integer
         Do While p_Dic.Exists(sKey & ".value" & IIf(i > 0, "" & i, ""))
           If i > 0 Then sList = sList & ","
           Select Case LCase(sPrefix)
             Case "s"
               sList = sList & "'" & Replace(p_Dic(sKey & ".value" & IIf(i > 0, "" & i, "")), "'", "''") & "'"
             Case "d"
               sList = sList & fncDateToSTR(p_Dic(sKey & ".value" & IIf(i > 0, "" & i, "")))
             Case "n"
               sList = sList & Replace("" & p_Dic(sKey & ".value" & IIf(i > 0, "" & i, "")), ",", ".")
           End Select
         Loop
         
         smid = " and " & IIf(idOperation = opNIN, "not ", "") & " in ( " & sList & ")"

        Else
         sOperand = p_Dic(sKey & ".value") ' с чем сравниваем, значение параметра
         
         If IsNull(sOperand) Then
           sOperand = "(null)"
         Else
            Select Case sPrefix
             Case "d"
              sOperand = fncDateToSTR(sOperand)
             Case "s"
              sOperand = "'" & sOperand & "'"
            End Select
         End If
         
         If (idOperation And opBTW) = opBTW Then '  для операторов between собираем второе значение
          sOperand2 = p_Dic(sKey & ".value1")
          
          If IsNull(sOperand2) Then
           sOperand2 = "(null)"
          Else
           Select Case sPrefix
            Case "d"
             sOperand2 = fncDateToSTR(sOperand2)
            Case "s"
             sOperand2 = "'" & Replace(sOperand2, "'", "''") & "'"
           End Select
          End If
          
          smid = " and " & sOperand & IIf((idOperation And 4096) = 0, " <= ", " < ") & SFMT
          smid = smid & " and " & SFMT & IIf((idOperation And 8192) = 0, " <= ", " < ") & sOperand2
         Else
          Select Case idOperation
           Case opEQ
            smid = " and " & SFMT & " = " & sOperand
           Case opNEQ
            smid = " and " & SFMT & " <> " & sOperand
           Case opGR
            smid = " and " & SFMT & " > " & sOperand
           Case opLS
            smid = " and " & SFMT & " < " & sOperand
           Case opNLS
            smid = " and " & SFMT & " >= " & sOperand
           Case opNGR
            smid = " and " & SFMT & " <= " & sOperand
           Case opcont
            smid = " and " & SFMT & " like '" & p_Dic("%") & Mid(sOperand, 2, Len(sOperand) - 2) & p_Dic("%") & "'"
           Case opSTART
            smid = " and " & SFMT & " like '" & Mid(sOperand, 2, Len(sOperand) - 2) & p_Dic("%") & "'"
           Case opNCont
            smid = " and not " & SFMT & " like '" & p_Dic("%") & Mid(sOperand, 2, Len(sOperand) - 2) & p_Dic("%") & "'"
          End Select
         End If
        End If
      End If
    ElseIf Left(LCase(SFMT), 7) = "sqldate" Then
     SFMT = Split(SFMT, ":")
     If IsNull(p_Dic(sKey)) Or IsEmpty(p_Dic(sKey)) Then
      If UBound(SFMT) > 0 Then
       smid = SFMT(1)
      Else
       smid = "(null)"
      End If
     Else
      smid = fncDateToSTR(p_Dic(sKey))
     End If
    ElseIf LCase(Left(sKey, 3)) = "fnc" Then
     smid = Application.Run(sKey, p_Dic(SFMT))
    ElseIf LCase(sKey) = "get" Then
     smid = Report.GetValue(SFMT & "", p_Dic, 1)
    Else
     smid = Format(p_Dic(sKey), SFMT)
    End If
   Else
    If p_Dic.Exists(Trim(smid)) Then
     smid = p_Dic(Trim(smid))
    
    Else
     On Error Resume Next
     sKey = Trim(smid)
     smid = Application.Eval(sKey)
     If Err Then
       smid = ""
       Err.Clear
     Else
       p_Dic(sKey) = smid
     End If
     On Error GoTo ErrorMet
 
    End If
   End If
   FilterFmt = Left(FilterFmt, p - 1) & smid & Mid(FilterFmt, pn + 1)
   p = p + Len(smid)
ErrorMet:
   If Err Then
    Err.Clear
    FilterFmt = Left(FilterFmt, p - 1) & "{Error}" & Mid(FilterFmt, pn + 1)
   End If
  End If
  Do While True
   p = InStr(p, FilterFmt, "%")
   If p = 0 Then Exit Do
   If Mid(FilterFmt, p + 1, 1) <> "%" Then Exit Do
   FilterFmt = Left(FilterFmt, p) & Mid(FilterFmt, p + 2)
   p = p + 1
  Loop
 Loop
End Function


Function InSet(spKey As Variant, ParamArray apArgs() As Variant) As Boolean
'Если первый параметр равен одному из последующих то возвращает Истину. В данной реализации Null = Null возвращает так же Истину
'#param spKey: Проверяемое значение
'#param apArgs: Одно или несколько тестовых значений. Если значение является массивом, то проверяется каждый элемент массива. Массивы могут быть вложенными
 Dim v As Variant
 v = apArgs
 InSet = InSetInner(spKey, v)
End Function

Function InSetInner(spKey As Variant, apArgs As Variant) As Boolean
 Dim bvIsNull As Boolean
 Dim i As Integer
 
 If IsEmpty(spKey) Then spKey = Null
 bvIsNull = IsNull(spKey)
 InSetInner = True
 For i = 0 To UBound(apArgs)
  If IsArray(apArgs(i)) Then
   If InSetInner(spKey, apArgs(i)) Then Exit Function
  ElseIf bvIsNull Then
   If IsNull(apArgs(i)) Then Exit Function
  ElseIf apArgs(i) = spKey Then
   Exit Function
  End If
 Next
 InSetInner = False
End Function


Public Function fncDateToSTR(dDate) As String
'Форматирует дату в литерал для использования в SQL запросе
'#param dDate: Дата
 If IsNull(dDate) Then
   fncDateToSTR = "NULL"
 Else
   fncDateToSTR = "#" & Format(dDate, "mm\/dd\/yyyy hh:nn:ss") & "#"
 End If
End Function



Public Function SelectOneValue(sql As String) As Variant
'Выполняет запрос и значение из первой колонки первой строки
'#param SQL: Текст запроса
 
 Dim rsdao
 Set rsdao = CurrentProject.Connection.Execute(sql)
 On Error GoTo noRecord
 If rsdao.EOF Then SelectOneValue = Empty Else SelectOneValue = rsdao.Fields(0).Value
 rsdao.Close
 Set rsdao = Nothing
 Exit Function
noRecord:
 SelectOneValue = Empty
 rsdao.Close
 Set rsdao = Nothing
End Function


'SUBBLOCK_BEGIN:BARCODE_COMMON
'Общие функции для формирования штрихкодов в форматe EMF.

Function longToByte(l)
  Dim tl: tl = l
  longToByte = Chr(tl Mod 256)
  tl = tl \ 256
  longToByte = longToByte & Chr(tl Mod 256)
  tl = tl \ 256
  longToByte = longToByte & Chr(tl Mod 256)
  tl = tl \ 256
  longToByte = longToByte & Chr(tl Mod 256)
End Function

Function intToByte(i)
  Dim ti: ti = i
  intToByte = Chr(ti Mod 256)
  ti = ti \ 256
  intToByte = intToByte & Chr(ti Mod 256)
End Function

Function block(fnc, data)
  block = intToByte(fnc) & data
  block = longToByte((Len(block) \ 2) + 2) & block
End Function

Function Point(x, y)
  Point = intToByte(x) & intToByte(y)
End Function

Function color(r, g, b)
  color = Chr(0) & Chr(b Mod 256) & Chr(g Mod 256) & Chr(r Mod 256)
End Function

Function RectAsPoligon(ByRef objCount, l, t, r, b)
  objCount = objCount + 1
  RectAsPoligon = block(&H324, intToByte(4) & Point(l, b) & Point(l, t) & Point(r, t) & Point(r, b))
End Function

Function CreatePenIndirect(ByRef objCount, PenStyle, pPoint, pColor)
  objCount = objCount + 1
  CreatePenIndirect = block(&H2FA, intToByte(PenStyle) & pPoint & pColor)
End Function

Function SelectObject(nObject)
  SelectObject = block(&H12D, intToByte(nObject))
End Function

Function CreateBrushIndirect(objCount, style, color, hatch)
  objCount = objCount + 1
  CreateBrushIndirect = block(&H2FC, intToByte(style) & color & intToByte(hatch))
End Function

Sub addInArray(ByRef spArray, ByRef pItem)
'Добавляет значение в массив
'#param spArray: Массив
'#param pItem: Добавляемый элемент

  If Not IsArray(spArray) Then spArray = Array()
  ReDim Preserve spArray(UBound(spArray) + 1)
  spArray(UBound(spArray)) = pItem
End Sub

Function zebra2wmf(s, xFactor, yFactor, ByRef MaxWidth)
  Dim recs, objCount, i, l, largest, size
  recs = Empty
  objCount = 0
  addInArray recs, CreatePenIndirect(objCount, 0, Point(0, 0), color(255, 0, 0))
  addInArray recs, SelectObject(objCount - 1)
  addInArray recs, CreateBrushIndirect(objCount, 0, color(0, 0, 0), 4)
  addInArray recs, SelectObject(objCount - 1)
  i = 1
  Do While i <= Len(s)
    l = 1
    Do While True
      If i + l > Len(s) Then Exit Do
      If Mid(s, i, 1) <> Mid(s, i + l, 1) Then Exit Do
      l = l + 1
    Loop
    If Mid(s, i, 1) = "|" Then
      addInArray recs, RectAsPoligon(objCount, i * xFactor, 0, (i + l) * xFactor - 1, yFactor)
    ElseIf Mid(s, i, 1) = "." Then
      addInArray recs, RectAsPoligon(objCount, i * xFactor, yFactor, i * xFactor, yFactor)
    ElseIf Mid(s, i, 1) = "L" Then
      addInArray recs, RectAsPoligon(objCount, i * xFactor, 0, (i + l) * xFactor - 1, yFactor * 1.2)
    End If
    MaxWidth = (i + l) * xFactor
    i = i + l
  Loop
  addInArray recs, block(0, Empty) 'EOF
  zebra2wmf = Join(recs, "")
  largest = 0
  For Each i In recs
    If Len(i) / 2 > largest Then largest = Len(i) / 2
  Next
  size = Len(zebra2wmf) / 2 + 9
  zebra2wmf = intToByte(1) & _
     intToByte(9) & _
     intToByte(&H100) & _
     intToByte(size Mod &H10000) & _
     intToByte(size \ &H10000) & _
     intToByte(UBound(recs)) & _
     longToByte(largest) & _
     intToByte(0) & zebra2wmf
End Function


Function isNumber(s)
'Внутренняя функция для формирования штрих кода. Проверяет что переданная строка состоит только из чисел.
'#param s: Строка

  Dim i
  isNumber = True
  For i = 1 To Len(s)
    If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then
      isNumber = False
      Exit Function
    End If
  Next
End Function
'SUBBLOCK_END



'SUBBLOCK_BEGIN:BARCODE_CODE128
'SUBBLOCK_DEPENDENCE:BARCODE_COMMON
'Добавляет поддержку штрих кодов в формате CODE128

Public Function Code128(pParamList, aArg As Variant) As String
'REPORT_FUNCTION: Code128(Текст;Ширина;Высота)
'Вставляет в документ картинку с штрихкодом в формате CODE128. Штрих код должен состоять только из букв английского алфавита и цифр. Контрольное число добавляется автоматически в конец.
'#param Текст: Кодируемый текст
'#param Ширина: Целевая ширина штрихкода
'#param Высота: Целевая высота штрихкода
'#return: RTF блок
  Dim byteStorage() As Byte, BCWidth
  If aArg(0) <> "" Then
    byteStorage = StrConv(zebra2wmf(code128_zebra(aArg(0), 3), 2, 40, BCWidth), vbFromUnicode)
    Code128 = PictureDataToRTF(byteStorage, aArg(1), aArg(2))
  Else
    Code128 = Empty
  End If
End Function


Function code128_zebra(SourceString, return_type)
'Внутренняя функция для формирования штрих кода. Формирует штрих код в формате CODE128
'#param SourceString: Кодируемый текст
'#param return_type: Тип возвращаемого результата
' {*} 0 - Кодирует для вывода специальным шрифтом
' {*} 1 - Формат для чтения человеком
' {*} 2 - возвращает контрольню сумму
' {*} 3 - Возвращает в виде последовательности символов `|` и ` `


 Dim i, dataToFormat, n, currentEncoding, weightedTotal, checkDigitValue, stringlen, currentValue, dataToPrint
  
 If IsNull(SourceString) Then Exit Function
 If SourceString = "" Then Exit Function
 
  
 i = 1
 dataToFormat = Trim(SourceString)
 stringlen = Len(dataToFormat)

 

 If return_type = 1 Then
   'Просто форматируем в переданное значение
   i = 1
   code128_zebra = ""
   For i = 1 To stringlen
     n = Asc(Mid(dataToFormat, i, 1))
     If i < Len(dataToFormat) - 2 And n = 202 Then
       n = CLng(Mid(dataToFormat, i + 1, 2))
       If ((i < Len(dataToFormat) - 4) And ((n >= 80 And n <= 81) Or (n >= 31 And n <= 34))) Then
         code128_zebra = code128_zebra & " (" & Mid(dataToFormat, i + 1, 4) & ") "
         i = i + 4
       ElseIf ((i < Len(dataToFormat) - 3) And ((n >= 40 And n <= 49) Or (n >= 23 And n <= 25))) Then
         code128_zebra = code128_zebra & " (" & Mid(dataToFormat, i + 1, 3) & ") "
         i = i + 3
       ElseIf ((i < Len(dataToFormat) - 2) And ((n >= 0 And n <= 30) Or (n >= 90 And n <= 99))) Then
         code128_zebra = code128_zebra & " (" & Mid(dataToFormat, i + 1, 2) & ") "
         i = i + 2
       End If
     ElseIf Asc(Mid(dataToFormat, i, 1)) < 32 Then
       code128_zebra = code128_zebra & " "
     ElseIf Asc(Mid(dataToFormat, i, 1)) > 31 And Asc(Mid(dataToFormat, i, 1)) < 128 Then
       code128_zebra = code128_zebra & Mid(dataToFormat, i, 1)
     End If
     i = i + 1
   Next
 Else
   n = Asc(Mid(dataToFormat, 1, 1))
   If n < 32 Then
     code128_zebra = Chr(203) 'A
     currentEncoding = "A"
   ElseIf (Len(dataToFormat) > 4 And isNumber(Mid(dataToFormat, 1, 4))) Or n = 202 Then
     code128_zebra = Chr(205) 'C
     currentEncoding = "C"
   ElseIf n >= 32 And n < 127 Then
     code128_zebra = Chr(204) 'B
     currentEncoding = "B"
   Else
     
   End If
     
   
   Do While i <= stringlen
     If Mid(dataToFormat, i, 1) = Chr(202) Then
       code128_zebra = code128_zebra & Chr(202)
        
     ElseIf ((i < stringlen - 2) And isNumber(Mid(dataToFormat, i, 4))) Or ((i < stringlen) And isNumber(Mid(dataToFormat, i, 2)) And (currentEncoding = "C")) Then
   
       If currentEncoding <> "C" Then code128_zebra = code128_zebra & Chr(199)
       currentEncoding = "C"
       currentValue = CLng(Mid(dataToFormat, i, 2))
   
       If currentValue < 95 Then code128_zebra = code128_zebra & Chr(currentValue + 32) Else code128_zebra = code128_zebra & Chr(currentValue + 100)
       i = i + 1
   
     ElseIf ((Asc(Mid(dataToFormat, i, 1)) < 31) Or ((currentEncoding = "A") And (Asc(Mid(dataToFormat, i, 1)) > 32 And (Asc(Mid(dataToFormat, i, 1))) < 96))) Then
   
       If currentEncoding <> "A" Then code128_zebra = code128_zebra & Chr(201)
       currentEncoding = "A"
       n = Asc(Mid(dataToFormat, i, 1))
       If n = 32 Then code128_zebra = code128_zebra & Chr(194) Else If n < 32 Then code128_zebra = code128_zebra & Chr(n + 96) Else code128_zebra = code128_zebra & Chr(n)
   
     ElseIf ((Asc(Mid(dataToFormat, i, 1))) > 31 And (Asc(Mid(dataToFormat, i, 1))) < 127) Then
   
       If currentEncoding <> "B" Then code128_zebra = code128_zebra & Chr(200)
       currentEncoding = "B"
       n = Asc(Mid(dataToFormat, i, 1))
       If n = 32 Then code128_zebra = code128_zebra & Chr(194) Else code128_zebra = code128_zebra & Chr(n)
     End If
     i = i + 1
   Loop
   
   If code128_zebra = "" Then
     Exit Function
   End If
   
   checkDigitValue = Asc(Mid(code128_zebra, 1, 1)) - 100
   For i = 2 To Len(code128_zebra)
     n = Asc(Mid(code128_zebra, i, 1))
     If n = 194 Then currentValue = 0 Else If n < 135 Then currentValue = n - 32 Else currentValue = n - 100
     checkDigitValue = checkDigitValue + currentValue * (i - 1)
   Next
   
   checkDigitValue = checkDigitValue Mod 103
   
   If checkDigitValue >= 95 Then checkDigitValue = Chr(checkDigitValue + 100) Else If checkDigitValue = 0 Then checkDigitValue = Chr(194) Else checkDigitValue = Chr(checkDigitValue + 32)
   
   If return_type = 0 Or return_type = 3 Then
     code128_zebra = code128_zebra & checkDigitValue & Chr(206) 'End
     
     If return_type = 3 Then
       dataToPrint = code128_zebra
       code128_zebra = ""
       Dim zebraArr: zebraArr = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "|| ||  ||  ", _
                                      "||  || ||  ", "||  ||  || ", "|  |  ||   ", "|  |   ||  ", "|   |  ||  ", "|  ||  |   ", "|  ||   |  ", "|   ||  |  ", _
                                      "||  |  |   ", "||  |   |  ", "||   |  |  ", "| ||  |||  ", "|  || |||  ", "|  ||  ||| ", "| |||  ||  ", "|  ||| ||  ", _
                                      "|  |||  || ", "||  |||  | ", "||  | |||  ", "||  |  ||| ", "|| |||  |  ", "||  ||| |  ", "||| || ||| ", "||| |  ||  ", _
                                      "|||  | ||  ", "|||  |  || ", "||| ||  |  ", "|||  || |  ", "|||  ||  | ", "|| || ||   ", "|| ||   || ", "||   || || ", _
                                      "| |   ||   ", "|   | ||   ", "|   |   || ", "| ||   |   ", "|   || |   ", "|   ||   | ", "|| |   |   ", "||   | |   ", _
                                      "||   |   | ", "| || |||   ", "| ||   ||| ", "|   || ||| ", "| ||| ||   ", "| |||   || ", "|   ||| || ", "||| ||| || ", _
                                      "|| |   ||| ", "||   | ||| ", "|| ||| |   ", "|| |||   | ", "|| ||| ||| ", "||| | ||   ", "||| |   || ", "|||   | || ", _
                                      "||| || |   ", "||| ||   | ", "|||   || | ", "||| |||| | ", "||  |    | ", "||||   | | ", "| |  ||    ", "| |    ||  ", _
                                      "|  | ||    ", "|  |    || ", "|    | ||  ", "|    |  || ", "| ||  |    ", "| ||    |  ", "|  || |    ", "|  ||    | ", _
                                      "|    || |  ", "|    ||  | ", "||    |  | ", "||  | |    ", "|||| ||| | ", "||    | |  ", "|   |||| | ", "| |  ||||  ", _
                                      "|  | ||||  ", "|  |  |||| ", "| ||||  |  ", "|  |||| |  ", "|  ||||  | ", "|||| |  |  ", "||||  | |  ", "||||  |  | ", _
                                      "|| || |||| ", "|| |||| || ", "|||| || || ", "| | ||||   ", "| |   |||| ", "|   | |||| ", "", "", "", "", "", "", "", "", "", _
                                      "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
                                      "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "|| ||  ||  ", "| |||| |   ", "| ||||   | ", _
                                      "|||| | |   ", "|||| |   | ", "| ||| |||| ", "| |||| ||| ", "||| | |||| ", "|||| | ||| ", "|| |    |  ", "|| |  |    ", _
                                      "|| |  |||  ", "||   ||| | ||", "|| ||  ||  ")
       For i = 1 To Len(dataToPrint)
         code128_zebra = code128_zebra & zebraArr(Asc(Mid(dataToPrint, i, 1)))
       Next
     End If
   ElseIf return_type = 2 Then
     code128_zebra = checkDigitValue
   End If
 End If
End Function

'SUBBLOCK_END

'SUBBLOCK_BEGIN:BARCODE_EAN13
'SUBBLOCK_DEPENDENCE:BARCODE_COMMON
'Добавляет поддержку штрих кодов в формате EAN13

Public Function EAN13(pParamList, aArg As Variant) As String
'REPORT_FUNCTION: EAN13(Текст;Ширина;Высота)
'Вставляет в документ картинку с штрихкодом в формате EAN13. Штрих код должен быть не больше 13 цифр, при этом крайняя правая цифра - это контрольная сумма
'#param Текст: Кодируемый текст
'#param Ширина: Целевая ширина штрихкода
'#param Высота: Целевая высота штрихкода
'#return: RTF блок

  Dim byteStorage() As Byte, BCWidth
  If aArg(0) <> "" Then
  
    byteStorage = StrConv(zebra2wmf(EAN13_zebra(aArg(0), False), 2, 40, BCWidth), vbFromUnicode)
    EAN13 = PictureDataToRTF(byteStorage, aArg(1), aArg(2))
  Else
    EAN13 = Empty
  End If
End Function


Public Function EAN13CheckNumber(ByVal Code)
'Расчитывает контрольную сумму для штрихкода в формате EAN13
'#param Code: Число для кодировки
  
  Dim sCode, i, CheckSum
  sCode = Code & ""
  CheckSum = 0
  For i = 0 To Len(sCode) - 1
    If i Mod 2 = 0 Then EAN13CheckNumber = EAN13CheckNumber + CInt(Mid(sCode, Len(sCode) - i, 1)) * 3 Else EAN13CheckNumber = EAN13CheckNumber + CInt(Mid(sCode, Len(sCode) - i, 1))
  Next
  CheckSum = 10 - CheckSum Mod 10
  EAN13CheckNumber = CheckSum
End Function

Function EAN13_zebra(ByVal Code, addCheckSum)
' es - 31.07.2023
'
' -------------------------------------------------------------------------------------------------/
On Error GoTo EAN13_Err

  Dim sCode, zebra, codeSchema, i
  sCode = Code & ""
  
  If Not isNumber(sCode) Then Exit Function
  
  
  
  If addCheckSum Then Code = Code & EAN13CheckNumber(sCode)

  sCode = Right("0000000000000" & Code, 13)

  zebra = Array(Array("   || |", "|||  | ", " |  |||"), _
                Array("  ||  |", "||  || ", " ||  ||"), _
                Array("  |  ||", "|| ||  ", "  || ||"), _
                Array(" |||| |", "|    | ", " |    |"), _
                Array(" |   ||", "| |||  ", "  ||| |"), _
                Array(" ||   |", "|  ||| ", " |||  |"), _
                Array(" | ||||", "| |    ", "    | |"), _
                Array(" ||| ||", "|   |  ", "  |   |"), _
                Array(" || |||", "|  |   ", "   |  |"), _
                Array("   | ||", "||| |  ", "  | |||"))
                
  
  codeSchema = Array(Array(0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1), _
                     Array(0, 0, 2, 0, 2, 2, 1, 1, 1, 1, 1, 1), _
                     Array(0, 0, 2, 2, 0, 2, 1, 1, 1, 1, 1, 1), _
                     Array(0, 0, 2, 2, 2, 0, 1, 1, 1, 1, 1, 1), _
                     Array(0, 2, 0, 0, 2, 2, 1, 1, 1, 1, 1, 1), _
                     Array(0, 2, 2, 0, 0, 2, 1, 1, 1, 1, 1, 1), _
                     Array(0, 2, 2, 2, 0, 0, 1, 1, 1, 1, 1, 1), _
                     Array(0, 2, 0, 2, 0, 2, 1, 1, 1, 1, 1, 1), _
                     Array(0, 2, 0, 2, 2, 0, 1, 1, 1, 1, 1, 1), _
                     Array(0, 2, 2, 0, 2, 0, 1, 1, 1, 1, 1, 1))(CInt(Mid(sCode, 1, 1)))
  
  EAN13_zebra = ".        L L"
  
  For i = 2 To 13
    EAN13_zebra = EAN13_zebra & zebra(CInt(Mid(sCode, i, 1)))(codeSchema(i - 2))
    If i = 7 Then EAN13_zebra = EAN13_zebra & " L L "
  Next
  
  EAN13_zebra = EAN13_zebra & "L L        ."
  ' -------------------------------------------------------------------------------------------------/
EAN13_End:
    On Error Resume Next
    Err.Clear
    Exit Function
' -------------------------------------------------------------------------------------------------/
EAN13_Err:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function :" & _
        "EAN13 - Report.", vbCritical, "Error!"
    Err.Clear
    EAN13_zebra = Empty
    Resume EAN13_End
End Function

'SUBBLOCK_END
