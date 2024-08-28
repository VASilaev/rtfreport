Option Compare Text
Option Explicit
'@Ignore ProcedureNotUsed

' Модуль генерации отчетов в формате RTF из шаблона
' Версия 2.0. 2024 год
' Больше информации на странице https://github.com/VASilaev/rtfreport

Const cReportTable As String = "t_rep"
Const StopSym As String = "(;,+-*/&^%<>= )"

Private GlobalContext As Object

Public Sub InstallRepSystem()
  'Создает необходимые таблицы. Для хранения шаблонов внутри таблицы
  If Not IsHasRepTable() Then
    With CurrentDb()
      .Execute "CREATE TABLE " & cReportTable & " " _
             & "(id counter CONSTRAINT PK_rep PRIMARY KEY, " _
             & "sCaption CHAR(255), sOrignTemplate memo, " _
             & "dEditTemplate date, sDescription char(255),clTemplate memo);"
      .TableDefs.Refresh
    End With
  End If
End Sub

'@EntryPoint
Public Sub InstallReportTemplate()
  'Добавляет файл в хранилище шаблонов, после чего сам файл можно удалить, а шаблон вызывать по коду или имени
  InstallRepSystem

  Dim dlgOpenFile, sFileName As String, idReport, sFilePath As String, sCurPath As String
  
#If VBA7 Then
  
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
  
#Else
  
  
    sFilePath = ahtCommonFileOpenSave( _
                Filter:=ahtAddFilterItem(vbNullString, "RTF шаблон", "*.rtf"), OpenFile:=True, _
                DialogTitle:="Выберите шаблон", _
                Flags:=ahtOFN_HIDEREADONLY)
  
#End If

  sFileName = GetFile(sFilePath)
  sFileName = Mid(sFileName, 1, Len(sFileName) - Len(GetExt(sFileName)))

  idReport = SelectOneValue("select id from " & cReportTable & " where ucase(sCaption) = '" & sFileName & "'")
  If IsEmpty(idReport) Then
    sCurPath = GetPath(CurrentDb.Name)
    If UCase(Left(sFilePath, Len(sCurPath))) = UCase(sCurPath) Then sFilePath = "." & Mid(sFilePath, Len(sCurPath))
    CurrentDb().Execute "insert into " & cReportTable & " (sCaption, sOrignTemplate) values ('" & Replace(sFileName, "'", "''") & "','" & Replace(sFilePath, "'", "''") & "');"
    idReport = SelectOneValue("select id from " & cReportTable & " where ucase(sCaption) = '" & UCase(sFileName) & "'")
    GetTemplate CLng(idReport)
    MsgBox "Отчет с именем """ & sFileName & """ зарегистрирован с кодом " & idReport
  Else
    MsgBox "Отчет с именем """ & sFileName & """ уже существует по кодом " & idReport
  End If
End Sub

Private Function IsHasRepTable()
  'Проверка наличия в БД таблицы с сохраненными отчетами, имя таблицы задается константой `cReportTable`

  Dim vDB As Object
  IsHasRepTable = True
  Set vDB = CurrentDb()
  On Error GoTo NoTable
  With vDB.TableDefs(cReportTable): End With
  On Error GoTo 0
  Exit Function
NoTable:
  On Error GoTo 0
  IsHasRepTable = False
End Function

Public Sub PrintReport(ByRef vReport, Optional ByRef dic As Object, Optional sFile As String = vbNullString, Optional bPrint As Boolean = False)
  'Запускает формирование документа из шаблона
  '#param vReport: Идентификатор шаблона.
  'Если число, то ищется в таблице с сохраненными отчетами, в противном случае считается что это имя файла.
  'Для поиска относительно местоположения БД используйте в начале `.\`.
  'Если такого файла не существует, то шаблон ищется по заголовку (`sCaption`) в таблице с сохраненными отчетами.
  '#param dic: Словарь с окружением, можно передать nothing если явных входных параметров нет
  '#param sFile: Имя выходного файла, если его не указать то будет создан во временной папке с именем tmp_n где n - порядковый номер
  '#param bPrint: Отправить документ на печать

  Dim fso
  Dim tf                                         'As TextStream
  Dim asTemplate, i As Variant
  Dim sPathOrig As String, sExtension As String

  Set fso = CreateObject("scripting.FileSystemObject")
 
  If dic Is Nothing Then
    If GlobalContext Is Nothing Then
      Set GlobalContext = CreateObject("Scripting.Dictionary")
      GlobalContext.CompareMode = 1
    End If
    Set dic = GlobalContext
  End If

  If IsNumeric(vReport) Then
    If IsHasRepTable() Then
      asTemplate = KRNReport.GetTemplate(CLng(vReport))
    Else
      Err.Raise 1000, , "Не найден шаблон """ & vReport & """"
    End If
  Else
  
    sPathOrig = vReport
    If Left(sPathOrig, 2) = ".\" Then sPathOrig = GetPath(CurrentDb.Name) & Mid(sPathOrig, 3)
   
    sExtension = LCase(fso.GetExtensionName(sPathOrig))
   
    If fso.FileExists(sPathOrig) Then
      Select Case sExtension
      Case "rtf"
        asTemplate = Array("rtf", PrepareRTF(sPathOrig))
      End Select
    ElseIf IsHasRepTable() Then
      i = Empty
      If IsHasRepTable() Then i = SelectOneValue("select id from " & cReportTable & " where ucase(sCaption) = '" & UCase(vReport) & "'")
      If IsEmpty(i) Then
        Err.Raise 1000, , "Не найден шаблон """ & vReport & """"
      Else
        asTemplate = KRNReport.GetTemplate(CLng(i))
      End If
    Else
      Err.Raise 1000, , "Не найден шаблон """ & vReport & """"
    End If
  End If
 
  dic("Extension") = asTemplate(0)
  dic("Date") = Date
  dic("Now") = Now
  
  If sFile = vbNullString Then
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
  KRNReport.MakeReport svTemplate, tf, dic
  tf.Close
  Set tf = Nothing
  Set fso = Nothing
 
  If InSet(asTemplate(0), "rtf") Then
    If bPrint Then
      CreateObject("WScript.Shell").Run "winword """ & sFile & """ /q /n /mFilePrintDefault /mFileSave /mFileExit", vbNormalFocus
    Else
      CreateObject("WScript.Shell").Run "winword """ & sFile & """", vbNormalFocus
    End If
  ElseIf InSet(asTemplate(0), "txt") Then
    CreateObject("WScript.Shell").Run "notepad """ & sFile & """", vbNormalFocus
  End If
End Sub

Function BuildParam(ByRef pDic As Object, ParamArray pdata())
  'Обновляет в контексте переменные
  '`BuildParam(pDic, Key, Value [, Key, Value])`
  '#param pDic: Текст
  '#param Key: Имя переменной
  '#param Value: Значение переменной. Объекты должны быть завернуты в массив: `array(MyObject)`

  Dim i, vData
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
    If IsArray(vData(i)) Then
      BuildParam pDic, Null, vData(i)
      i = i + 1
    ElseIf i < UBound(vData) Then
      If IsNull(vData(0)) And IsArray(vData(1)) Then
        BuildParam pDic, Null, vData(i + 1)
      Else
        pDic(vData(i) & vbNullString) = vData(i + 1)
      End If
      i = i + 2
    Else
      Err.Raise 1000, , "Не парное число параметров!"
    End If
    
  Loop
  Set BuildParam = pDic
End Function

Function LPad(ByRef s As String, ByVal ch As String, ByVal TotalCnt As Integer) As String
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

Function GetRegExp(ByVal spPattern As Variant) As Variant
  'Создает объект RegExpс заданным паттерном
  'param spPattern: Паттерн регулярного выражения. в начале могут быть дополнительные модификаторы:
  ' {*} \g - Глолбальный поиск
  ' {*} \i - Поиск без учета регистра
  ' {*} \m - Символ переводла сроки заканичвает текущую строкуи начинает новую. Влияет на модификатор ^ и $
  Dim svPattern: svPattern = spPattern: Set GetRegExp = CreateObject("VBScript.RegExp")
  GetRegExp.Global = False: GetRegExp.IgnoreCase = False: GetRegExp.Multiline = False
  Do While Left(svPattern, 1) = "\"
    Select Case LCase(Mid(svPattern, 2, 1))
    Case "g": GetRegExp.Global = True
    Case "i": GetRegExp.IgnoreCase = True
    Case "m": GetRegExp.Multiline = True
    Case Else: Exit Do
    End Select
    svPattern = Mid(svPattern, 3)
  Loop
  GetRegExp.Pattern = svPattern
End Function

Private Function GetEscape(ByRef sBuf As String, ByRef iPOS As Long) As String
  'Внутренняя функция разбора RTF. Возвращает экрнаированное значение за символом `\`
  '#param sBuf: Буфер
  '#param iPOS: Текущая позиция в буфере. Позиция обновляется.

  If Mid(sBuf, iPOS, 1) = "\" Then
    Select Case Mid(sBuf, iPOS + 1, 1)           '= "\"
    Case "'"
      GetEscape = Chr(CInt("&h" & Mid(sBuf, iPOS + 2, 2)))
      'iPOS = iPOS + 4
    Case "\", "{", "}"
      GetEscape = Mid(sBuf, iPOS + 1, 1)
      'iPOS = iPOS + 2
    Case Else
      GetEscape = vbNullString
    End Select
  Else
    GetEscape = vbNullString
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

  Dim State As Integer, ch As String
  
  '0, 10 - start

  '50 - text
  '100 - control
  
  GetToken = vbNullString
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
      If ch = vbNullString Then
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
  SkipBlock = vbNullString
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


  Dim sOpt As String, State As Long, ch As String, txt As String, sTmpToken As String
  
  
  State = 0
  ch = GetToken(sBuf, iPOS)
  
  Do While ch <> "cEOF"
    Select Case State
    Case 0
      Select Case LCase(ch)
      Case "c\flddirty", "c\fldedit", "c\fldlock", "c\fldpriv"
        State = State 'DoNothing
      Case "c{"
        State = 1                                ' ждем \*
      End Select
    Case 1
      If ch = "c\*" Then
        State = 2                                'ждем \fldinst
      Else
        Err.Raise 2001
      End If
    Case 2
      If ch = "c\fldinst" Then
        State = 3                                ' ждем первого кортежа с текстом
      Else
Debug.Print iPOS
        Err.Raise 1002
      End If
    Case 3, 13
    
      Select Case LCase(ch)
      Case "tref"
        State = State 'DoNothing
      Case "c{"
        ch = GetToken(sBuf, iPOS)
        If State = 3 Then sOpt = vbNullString
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
            State = 13                           'параметры берутся только из первого заполненого картежа
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

Function RemoveTag(ByRef rtf As String, ByVal tag As String) As String
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
    
    'Сдвинем курсор в случае если перевод строки, он не влияет не рендер
    Do While InSet(Mid(RemoveTag, ep, 1), vbCr, vbLf)
      ep = ep + 1
    Loop
 
    If Mid(rtf, tp - 1, 1) <> " " And Not InSet(Mid(RemoveTag, ep, 1), " ", "\", "{", "}") Then sSpace = " " Else sSpace = vbNullString

    RemoveTag = Mid(RemoveTag, 1, tp - 1) & sSpace & Mid(RemoveTag, ep)
    tp = InStr(1, RemoveTag, tag, vbTextCompare)
  Loop

End Function

Private Function RTF_SkipPar(ByRef sBuf As String, ByRef tp As Long) As String
  'Внутренняя функция разбора RTF. Пропускает перевод строки если он идет непосредственно за текстом.
  '#param sBuf: Буфер
  '#param tp: Текущая позиция разбора

  Dim LP As Long
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

Public Function InsertAddress(ByRef ts As String, adr As Long, iPOS As Variant) As String
  'Внутрення функция формирования шаблона. Вставляет адрес (8 символов в шестнадцетиричном виде) вместо заглушки
  '#param ts: Текущий буфер
  '#param adr: Вставляемый адрес
  '#param iPOS: Смещение в буфере куда нужно вставить адрес
  If VarType(iPOS) = vbString Then
    InsertAddress = ts
    
    If iPOS <> vbNullString Then
      Dim iPosition
      For Each iPosition In Split(iPOS, ",")
        InsertAddress = InsertAddress(InsertAddress, adr, CLng(iPosition))
      Next
    End If
  Else
    InsertAddress = Mid(ts, 1, iPOS - 1) & LPad(Hex(adr), "0", 8) & Mid(ts, iPOS + 8)
  End If
End Function

Public Function PrepareRTF(ByRef sFile As String) As String
  'Компилирует RTF файл в внутренний формат шаблона
  '#param sFile: Имя файла для разбора. Если строка начинается с `raw`, то предполагается что передано содержимое шаблона в формате rtf


  Dim ts As String
  Dim cp As Long, tp As Long
  Dim iStrucLevel As Long
  Dim aStrucStack(128) As Variant
  Dim nState, skipConst                          '/**Флаг пропуска куска RTF, в шаблон не включается*/
  Dim Res As String
  Dim MetList As String
  Dim sFmtTmp, sTXT, SFMT, sFnc, sOpt

  iStrucLevel = 0

  If Left(sFile, 3) = "raw" Then
    ts = Mid(sFile, 4)
  Else
    With CreateObject("scripting.FileSystemObject").OpenTextFile(sFile)
      ts = .ReadAll
      .Close
    End With
  End If
 
 
  cp = 1
  Res = "GOTO00000015        "                   'если будет список меток то заменится на call
 
  Dim re
  Set re = CreateObject("VBScript.RegExp")
  re.IgnoreCase = True
  re.Global = True
 
 'Удалим историю редактирования
  re.Pattern = "( | ?\r\n)?\\(pararsid|insrsid|charrsid|sectrsid|styrsid|tblrsid)\d+"
  ts = re.Replace(ts, vbNullString)
  
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
  ts = RemoveTag(ts, "{\info")
  ts = RemoveTag(ts, "{\sp{\sn metroBlob")

  re.Multiline = True
  re.IgnoreCase = True
  re.Global = False
  re.Pattern = "^\s*[_0-9а-яa-zё]+\s*\(.*\)$"
 
  nState = 0

  iStrucLevel = 1
  aStrucStack(iStrucLevel - 1) = Array(-1)
 
  skipConst = False

 
  Do While cp <= Len(ts)
 
    Select Case nState
    Case 0

      ' ищем данные до следующего Field
      tp = InStr(cp, ts, "{\field", vbTextCompare)
  
      If tp <> 0 Then
        ' у нас есть поле

        If Trim(Mid(ts, cp, tp - cp)) <> vbNullString And Not skipConst Then Res = Res & "PRNT" & LPad(Hex(tp - cp), "0", 8) & Mid(ts, cp, tp - cp)
        cp = tp
        tp = tp + 7
        sFmtTmp = ParseField(ts, tp)

        If LCase(Left(Trim(sFmtTmp(0)), 3)) = "ref" Then sTXT = Trim(Mid(Trim(sFmtTmp(0)), 4)) Else sTXT = Trim(sFmtTmp(0))
        SFMT = sFmtTmp(1)

        re.Pattern = "^\s*[_0-9а-яa-zё]+\s*\(.*\)$"
   
        If Not re.test(GetRegExp("\g[\r\n]+").Replace(sTXT, vbNullString)) Then
          'это не правильное поле оставляем его как есть
          If Trim(Mid(ts, cp, tp - cp)) <> vbNullString Then Res = Res & "PRNT" & LPad(Hex(tp - cp), "0", 8) & Mid(ts, cp, tp - cp)
        Else
          sFnc = Mid(sTXT, 1, InStr(1, sTXT, "(") - 1)
          sOpt = Mid(sTXT, InStr(1, sTXT, "(") + 1)
          'If Right(sOpt, 1) <> ")" Then MsgBox "Ожидается в конце ')' в выражении " & stxt
          sOpt = Trim(Left(sOpt, Len(sOpt) - 1))


          nState = 2
          skipConst = False
    
          Select Case LCase(Trim(sFnc))
          Case "scan"
            If sOpt = vbNullString Then Err.Raise 1020, , "Ожидается выражение после scan"
            Dim nForPos, svCursorName, svSQLBody: nForPos = InStr(LCase(sOpt), " for ")
            If nForPos = 0 Then Err.Raise 1021, , "Параметр scan должен быть вида <Имя курсора> for <Выражение строкового вида>"
            svCursorName = Trim(Mid(sOpt, 1, nForPos - 1))
            svSQLBody = Mid(sOpt, nForPos + 5)
            Dim bNewPage
            bNewPage = LCase(Left(svCursorName, 8)) = "newpage "
            If bNewPage Then svCursorName = Trim(Mid(svCursorName, 9))
      
            Res = Res & "OPRS" & LPad(Hex(Len(svCursorName)), "0", 3) & svCursorName & LPad(Hex(Len(svSQLBody)), "0", 4) & svSQLBody & "________"
            aStrucStack(iStrucLevel) = Array(10, Len(Res) + 1, Str(Len(Res) - 7), 0, bNewPage) 'цикл, адрес старта цикла, адрес для вставки окончания цикла-11
            iStrucLevel = iStrucLevel + 1
            nState = 3                           'Если за полем идет \par то отбрасываем его
      
          Case "scanentry"
            If aStrucStack(iStrucLevel - 1)(0) <> 10 Then Err.Raise 8001, , "ScanEntry должен идти после Scan"
            If aStrucStack(iStrucLevel - 1)(3) > 0 Then Err.Raise 8001, , "ScanEntry должен быть " & IIf(aStrucStack(iStrucLevel - 1)(3) = 1, "один", "до ScanFooter")
      
            aStrucStack(iStrucLevel - 1)(1) = Len(Res) + 1
            aStrucStack(iStrucLevel - 1)(3) = 1
            nState = 3                           'Если за полем идет \par то отбрасываем его
      
          Case "scanfooter"
            If aStrucStack(iStrucLevel - 1)(0) <> 10 Then Err.Raise 8001, , "ScanFooter должен идти после Scan"
            If aStrucStack(iStrucLevel - 1)(3) > 1 Then Err.Raise 8001, , "ScanFooter должен быть один"
            aStrucStack(iStrucLevel - 1)(3) = 2
            nState = 3                           'Если за полем идет \par то отбрасываем его
     
            If aStrucStack(iStrucLevel - 1)(4) Then 'NewPage
              Res = Res & "NEXT" & LPad(Hex(Len(Res) + 24), "0", 8) & _
                    "GOTO" & LPad(Hex(Len(Res) + 54), "0", 8) & _
                    "PRNT00000007{\page}" & _
                    "GOTO" & LPad(Hex(aStrucStack(iStrucLevel - 1)(1)), "0", 8)
            Else
              Res = Res & "NEXT" & LPad(Hex(aStrucStack(iStrucLevel - 1)(1)), "0", 8)
            End If
     
          Case "endscan"
            iStrucLevel = iStrucLevel - 1
            If aStrucStack(iStrucLevel)(0) <> 10 Then Err.Raise 8001, , "Endscan нарушает структуру программы"
            If aStrucStack(iStrucLevel)(3) < 2 Then
              If aStrucStack(iStrucLevel)(4) Then 'NewPage
                Res = Res & "NEXT" & LPad(Hex(Len(Res) + 24), "0", 8) & _
                      "GOTO" & LPad(Hex(Len(Res) + 54), "0", 8) & _
                      "PRNT00000007{\page}" & _
                      "GOTO" & LPad(Hex(aStrucStack(iStrucLevel)(1)), "0", 8)
              Else
                Res = Res & "NEXT" & LPad(Hex(aStrucStack(iStrucLevel)(1)), "0", 8)
              End If
            End If
            Res = InsertAddress(Res, Len(Res) + 1, aStrucStack(iStrucLevel)(2))
               
          Case "next"
            Res = Res & "CONT" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
       
            'Пропустить часть шаблона, может забрать с собой перевод строки
          Case "skip"
            skipConst = True
          Case "endskip"
            If LCase(Trim(sOpt)) = "skip_lf" Then nState = 3
            skipConst = False
       
       
          Case "if"
            Res = Res & "JMPF" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt & "________"
      
            aStrucStack(iStrucLevel) = Array(0, Str(Len(Res) - 7), vbNullString) 'if then or elif
            iStrucLevel = iStrucLevel + 1
      
            'Если за полем идет \par то отбрасываем его
            nState = 3

      
          Case "elif"
            If Not aStrucStack(iStrucLevel - 1)(0) = 0 Then Err.Raise 8003, , "Elif нарушает структуру программы"
     
     
            Res = Res & "GOTO________"
            If aStrucStack(iStrucLevel - 1)(2) <> vbNullString Then aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) + ","
            aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) & Str(Len(Res) - 7)
      
            Res = InsertAddress(Res, Len(Res) + 1, CLng(aStrucStack(iStrucLevel - 1)(1)))
      
            Res = Res & "JMPF" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt & "________"
            aStrucStack(iStrucLevel - 1)(1) = Str(Len(Res) - 7)

            'Если за полем идет \par то отбрасываем его
            nState = 3

      
          Case "else"
            If Not aStrucStack(iStrucLevel - 1)(0) = 0 Then Err.Raise 8003, , "Else нарушает структуру программы"
            aStrucStack(iStrucLevel - 1)(0) = 1
      
            Res = Res & "GOTO"
            If aStrucStack(iStrucLevel - 1)(2) <> vbNullString Then aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) + ","
            aStrucStack(iStrucLevel - 1)(2) = aStrucStack(iStrucLevel - 1)(2) & Str(Len(Res) + 1)
            Res = Res & "________"
      
            Res = InsertAddress(Res, Len(Res) + 1, CLng(aStrucStack(iStrucLevel - 1)(1)))
            aStrucStack(iStrucLevel - 1)(1) = vbNullString
      
            'Если за полем идет \par то отбрасываем его
            nState = 3
      
          Case "endif"
            If Not InSet(aStrucStack(iStrucLevel - 1)(0), 0, 1) Then Err.Raise 8003, , "Endif нарушает структуру программы"
      
            iStrucLevel = iStrucLevel - 1
            If aStrucStack(iStrucLevel)(1) <> vbNullString Then
              If aStrucStack(iStrucLevel)(2) <> vbNullString Then aStrucStack(iStrucLevel)(2) = aStrucStack(iStrucLevel)(1) & ","
              aStrucStack(iStrucLevel)(2) = aStrucStack(iStrucLevel)(2) & aStrucStack(iStrucLevel)(1)
            End If
      
            Res = InsertAddress(Res, Len(Res) + 1, aStrucStack(iStrucLevel)(2))

            'Обрабатываем опциональный параметр пропуска перевода строки
            If LCase(Trim(sOpt)) = "skip_lf" Then nState = 3
      
         
          Case "calc"
            Res = Res & "CALC" & LPad(Hex(Len(sTXT)), "0", 3) & sTXT
          Case "sum"
            sTXT = "sum(" & sTXT & ")"
            Res = Res & "CALC" & LPad(Hex(Len(sTXT)), "0", 3) & sTXT
          Case "inc"
            sTXT = "inc(" & sTXT & ")"
            Res = Res & "CALC" & LPad(Hex(Len(sTXT)), "0", 3) & sTXT
          Case "cts"
            sTXT = "cts(" & sTXT & ")"
            Res = Res & "CALC" & LPad(Hex(Len(sTXT)), "0", 3) & sTXT
          Case "clr"
            sTXT = "clr(" & sTXT & ")"
            Res = Res & "CALC" & LPad(Hex(Len(sTXT)), "0", 3) & sTXT
          Case "fld", "f"
            If SFMT = "\no" Then
              Res = Res & "PRVL" & LPad(Hex(Len(sOpt)), "0", 3) & sOpt
            Else
              skipConst = False
              SFMT = "{" & Trim(SFMT)            '& "  "
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
            Err.Raise 1000, vbNullString, "Не известная комманда: " & sFnc
          End Select
        End If
        cp = tp
      Else
        tp = Len(ts) + 1
        Res = Res & "PRNT" & LPad(Hex(tp - cp), "0", 8) & Mid(ts, cp, tp - cp)
        cp = Len(ts) + 1
      End If
  
  
  
    Case 3
      Dim svToken, nvLength
      tp = cp
    
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
      tp = cp
      svToken = RTF_SkipPar(ts, tp)
      nState = 1
    End If
  
    If nState = 1 Then
      If svToken = "c\field" Then
        cp = InStr(cp, ts, "{\field", vbTextCompare)
      End If
      nState = 0
    End If

  
  Loop
 
  Res = Res & "ENDT"
  
  If MetList <> vbNullString Then
    Res = "CALL00d" & LPad(Trim(Len(Res) & vbNullString), "0", 13) & Mid(Res, 21) & MetList & "RETC"
  End If
 
  If iStrucLevel > 1 Then
    MsgBox IIf(aStrucStack(iStrucLevel - 1)(0) <> 10, "В ходе разбора файла обнаружен не закрытый блок IF. ", vbNullString) & IIf(aStrucStack(iStrucLevel - 1)(0) = 10, "В ходе разбора файла обнаружен не закрытый блок SCAN.", vbNullString)
    Exit Function
  End If

  PrepareRTF = Res
 
End Function

Private Function ToBool(ByRef Value As Variant, Optional ByVal bRaise As Boolean = True) As Variant
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
  If ParamList Is Nothing Then
    DumpContext = DumpContext & " Отсутствует"
  Else
    On Error Resume Next
    For Each key In ParamList.Keys
      If Not InSet(LCase(key), "%", "now", "date") Then
        If Not IsArray(ParamList(key)) Then
          DumpContext = DumpContext & vbCrLf & key & " = " & Mid(ParamList(key), 1, 128) & IIf(Len(ParamList(key)) > 128, "...", vbNullString)
        End If
      End If
    Next
    On Error GoTo 0
  End If

End Function

Public Function GetExpression(ByRef Formula As Variant, Optional ByRef ParamList As Variant = Nothing, Optional ByRef StartPos As Long = 1)
  'Внутренняя функция для формирования отчета. Разбирает варыжение и возвращает его значение. Строки задаются в двойных ковычках.
  'Значения переменных должны находиться в текущем контексте или находиться при помощи функции eval. Стартовые пробелы пропускаются.
  'Для получения более подробной информации см. справку.
  '#param Formula: Текст выражения
  '#param ParamList: Текущий контекст
  '#param StartPos: Текущее смещение в выражении

  Dim EvalExpression, ch, ch2

  Do While Mid(Formula, StartPos, 1) = " "
    StartPos = StartPos + 1
  Loop

  If Mid(Formula, StartPos, 1) = "(" Then
    StartPos = StartPos + 1
    GetExpression = GetExpression(Formula, ParamList, StartPos)
    Do While Mid(Formula, StartPos, 1) = " "
      StartPos = StartPos + 1
    Loop
  
    If Mid(Formula, StartPos, 1) = ")" Then
      StartPos = StartPos + 1
    Else
      Err.Raise 1000, , "Ожидается разделитель )"
    End If
  Else

    GetExpression = GetNotValue(Formula, ParamList, StartPos)
  
  End If

  Do While True
    Do While Mid(Formula, StartPos, 1) = " "
      StartPos = StartPos + 1
    Loop
  
    Dim tmpPos: tmpPos = StartPos
  
    ch = Mid(Formula, tmpPos, 1)
    ch2 = Mid(Formula, tmpPos + 1, 1)
    If ch = "<" And (ch2 = ">" Or ch2 = "=") Then
      tmpPos = tmpPos + 1
      ch = ch & ch2
    ElseIf ch = ">" And ch2 = "=" Then
      tmpPos = tmpPos + 1
      ch = ch & ch2
    ElseIf LCase(ch) >= "a" And LCase(ch) <= "z" Then
      Do While True
        ch2 = LCase(Mid(Formula, tmpPos + 1, 1))
        If Not (ch2 >= "a" And ch2 <= "z") Then Exit Do
        tmpPos = tmpPos + 1
        ch = ch & ch2
      Loop
    End If
    
    If ch = "&" And LCase(Mid(Formula, tmpPos + 1, 1)) = "h" Then ch = "&h"
    
    If Not InSet(ch, "and", "or", "xor", "like", "between", "mod", "+", "-", "*", "/", "<", "<=", "<>", ">", ">=", "^", "\", "&", "=") Then Exit Do
    
    If IsEmpty(EvalExpression) Then EvalExpression = ToSQL(GetExpression)
    
    EvalExpression = EvalExpression & " " & ch & " "
    
    StartPos = tmpPos + 1
    
    Do While Mid(Formula, StartPos, 1) = " "
      StartPos = StartPos + 1
    Loop
    
    If Mid(Formula, StartPos, 1) = "(" Then
      EvalExpression = EvalExpression & ToSQL(GetExpression(Formula, ParamList, StartPos))
    Else
      EvalExpression = EvalExpression & ToSQL(GetNotValue(Formula, ParamList, StartPos))
    End If
  Loop
  If Not IsEmpty(EvalExpression) Then GetExpression = Eval(EvalExpression)
End Function

Private Function GetNotValue(Formula As Variant, ByRef ParamList As Variant, ByRef StartPos As Long) As Variant
  'Внутренняя функция для формирования отчета. Разбирает один атом варыжения перед котором может быть ключевое слово Not и возвращает его значение. Строки задаются в двойных ковычках.
  'Значения переменных должны находиться в текущем контексте или находиться при помощи функции eval. Стартовые пробелы пропускаются.
  'Для получения более подробной информации см. справку.
  '#param Formula: Текст выражения
  '#param ParamList: Текущий контекст
  '#param StartPos: Текущее смещение в выражении

  Dim ch, ch2, tmpPos


  Do While True
    ch2 = LCase(Mid(Formula, StartPos, 1))
    If Not Mid(Formula, StartPos, 1) = " " Then Exit Do
    StartPos = StartPos + 1
  Loop


  If ch2 = "-" Then
    StartPos = StartPos + 1
    GetNotValue = -GetNotValue(Formula, ParamList, StartPos)
  ElseIf ch2 = "+" Then
    StartPos = StartPos + 1
    GetNotValue = GetNotValue(Formula, ParamList, StartPos)
  Else
    tmpPos = StartPos
    ch = vbNullString
    Do While True
      If Not (ch2 >= "a" And ch2 <= "z") Or tmpPos - StartPos > 3 Then Exit Do
      tmpPos = tmpPos + 1
      ch = ch & ch2
      ch2 = LCase(Mid(Formula, tmpPos, 1))
    Loop
        
    If ch = "not" Then
      StartPos = StartPos + 3
      GetNotValue = Not GetNotValue(Formula, ParamList, StartPos)
    ElseIf Mid(Formula, StartPos, 1) = "(" Then
      GetNotValue = GetExpression(Formula, ParamList, StartPos)
    Else
      GetNotValue = GetValue(Formula, ParamList, StartPos)
    End If
  End If
End Function

Public Function GetPartFromFormula(Formula As Variant, ByRef StartPos As Long)
  'список терминаторов
  Dim cp
  GetPartFromFormula = vbNullString
  Do While Mid(Formula, StartPos, 1) = " ": StartPos = StartPos + 1: Loop
  Do While StartPos <= Len(Formula)
    If LCase(Mid(Formula, StartPos, 2)) = "&h" Then
      GetPartFromFormula = GetPartFromFormula + Mid(Formula, StartPos, 1)
      StartPos = StartPos + 1
    ElseIf InStr(1, StopSym, Mid(Formula, StartPos, 1)) > 0 Then
      Exit Do
    ElseIf Mid(Formula, StartPos, 1) = "[" Then
      cp = InStr(StartPos, Formula, "]")
      If cp > 0 Then
        GetPartFromFormula = GetPartFromFormula + Mid(Formula, StartPos, cp - StartPos)
        StartPos = cp
      End If
    End If
    GetPartFromFormula = GetPartFromFormula + Mid(Formula, StartPos, 1)
    StartPos = StartPos + 1
  Loop
  Do While Mid(Formula, StartPos, 1) = " ": StartPos = StartPos + 1: Loop
End Function

Public Function GetValue(Formula As Variant, Optional ByRef ParamList As Variant = Nothing, Optional ByRef StartPos As Long = 1) As Variant
  'Внутренняя функция для формирования отчета. Разбирает один атом варыжения и возвращает его значение. Строки задаются в двойных ковычках.
  'Значения переменных должны находиться в текущем контексте или находиться при помощи функции eval. Стартовые пробелы пропускаются.
  'Для получения более подробной информации см. справку.
  '#param Formula: Текст выражения
  '#param ParamList: Текущий контекст
  '#param StartPos: Текущее смещение в выражении
  Dim sErrorMsg As String
  Dim cp As Long
  Dim sFnc As String
  Dim aArg(16) As Variant
  Dim nvSP As Long
  Dim objXML, objDocElem
  Dim byteStorage() As Byte
  Dim sCurrentChar As String

  If ParamList Is Nothing Then
    If GlobalContext Is Nothing Then
      Set GlobalContext = CreateObject("Scripting.Dictionary")
      GlobalContext.CompareMode = 1
    End If
    Set ParamList = GlobalContext
  End If

  'Пропускаем не значищие пробелы
  Do While Mid(Formula, StartPos, 1) = " "
    StartPos = StartPos + 1
  Loop

  nvSP = StartPos
  On Error GoTo OnError

  sCurrentChar = Mid(Formula, StartPos, 1)

  If sCurrentChar = """" Or sCurrentChar = "'" Or sCurrentChar = "`" Then 'строковая константа
  StartPos = StartPos + 1
  Do While StartPos <= Len(Formula)
    If Mid(Formula, StartPos, 2) = sCurrentChar & sCurrentChar Then
      sFnc = sFnc & sCurrentChar
      StartPos = StartPos + 2
    ElseIf Mid(Formula, StartPos, 1) = sCurrentChar Then
      GetValue = sFnc
      StartPos = StartPos + 1
      Exit Do
    Else
      sFnc = sFnc + Mid(Formula, StartPos, 1)
      StartPos = StartPos + 1
    End If
  Loop
ElseIf sCurrentChar = "#" Then                   'константа дата
  cp = InStr(StartPos + 1, Formula, "#")
  If cp > 0 Then
    sFnc = Mid(Formula, StartPos + 1, cp - StartPos - 1)
    StartPos = cp + 1
    GetValue = CDate(sFnc)
  Else
    Err.Raise 1000, , "Не удалось найти окончание константы даты"
  End If

ElseIf sCurrentChar = ")" Or StartPos > Len(Formula) Then
  GetValue = Null
  Exit Function
ElseIf sCurrentChar >= "0" And sCurrentChar <= "9" Or sCurrentChar = "." Then
  'Числовой литерал
  Dim bHasFrac As Boolean
  If sCurrentChar <> "." Then
    sFnc = sCurrentChar
    Do While True
      StartPos = StartPos + 1
      sCurrentChar = Mid(Formula, StartPos, 1)
      If sCurrentChar >= "0" And sCurrentChar <= "9" Then sFnc = sFnc & sCurrentChar Else Exit Do
    Loop
  End If
      
  bHasFrac = False
  If sCurrentChar = "." Then
    'дробная часть
    bHasFrac = True
    sFnc = sFnc & ","
    Do While True
      StartPos = StartPos + 1
      sCurrentChar = Mid(Formula, StartPos, 1)
      If sCurrentChar >= "0" And sCurrentChar <= "9" Then sFnc = sFnc & sCurrentChar Else Exit Do
    Loop
  End If
      
  If sCurrentChar = "@" Then
    StartPos = StartPos + 1
    GetValue = CDec(sFnc)
  ElseIf sCurrentChar = "#" Then
    StartPos = StartPos + 1
    GetValue = CDbl(sFnc)
  ElseIf sCurrentChar = "!" Then
    StartPos = StartPos + 1
    GetValue = CSng(sFnc)
  ElseIf bHasFrac Then
    GetValue = CDbl(sFnc)
  Else
    If sCurrentChar = "%" And Not bHasFrac Then
      StartPos = StartPos + 1
      GetValue = CInt(sFnc)
    ElseIf sCurrentChar = "&" And Not bHasFrac Then
      StartPos = StartPos + 1
      GetValue = CLng(sFnc)
    Else
      GetValue = CLng(sFnc)
    End If
  End If
        
      
Else
  sFnc = vbNullString
  sFnc = Trim(GetPartFromFormula(Formula, StartPos))
   
  If sFnc = vbNullString Then Err.Raise 1000, , "Ожидается какое либо значение."
   
  If Mid(Formula, StartPos, 1) = "(" Then        'Функция
    aArg(0) = 0
    StartPos = StartPos + 1                      'пропускаем (
    Do While StartPos <= Len(Formula)            'Собираем аргументы
      aArg(0) = aArg(0) + 1
      aArg(aArg(0)) = GetExpression(Formula, ParamList, StartPos)
      
      sCurrentChar = Mid(Formula, StartPos, 1)
      Do While sCurrentChar = " "
        StartPos = StartPos + 1
        sCurrentChar = Mid(Formula, StartPos, 1)
      Loop
  
      If sCurrentChar = ")" Then
        StartPos = StartPos + 1
        Exit Do
      ElseIf (Mid(Formula, StartPos, 1) = ";" Or Mid(Formula, StartPos, 1) = ",") Then
        StartPos = StartPos + 1                  ' пропускаем точку с запятой
      Else
        Err.Raise 1000, , "Ожидается разделитель ; или ,"
      End If
    Loop
      
      
  
    Select Case LCase(sFnc)

    Case "open"
      Dim objStream, fso
      Set fso = CreateObject("scripting.FileSystemObject")
      
      If aArg(0) = 0 Then
        GetValue = "Требуется имя файла"
      Else
        If Left(aArg(1), 2) = ".\" Then aArg(1) = GetPath(CurrentDb.Name) & Mid(aArg(1), 3)
      
        If Not fso.FileExists(aArg(1)) Then
          GetValue = "Файл '" & aArg(1) & "' не найден."
        Else
        
          Set objStream = CreateObject("ADODB.Stream")
          If aArg(2) & vbNullString <> vbNullString Then
            objStream.Charset = aArg(2)
            objStream.Type = 2                     'adTypeText
          Else
            objStream.Type = 1                     ' adTypeBinary
          End If
        
          objStream.Open
        
          objStream.LoadFromFile (aArg(1))
          GetValue = objStream.Read()
          Set objStream = Nothing
        End If
      End If
      Set fso = Nothing
    Case "attach"                                ' 1 параметр - это поле attachment, второе поле - это маска поиска
      Dim objRegExp As Object, FileName As Variant
    
      Set objRegExp = CreateObject("VBScript.RegExp")
      objRegExp.Global = False
      objRegExp.IgnoreCase = True
      objRegExp.Multiline = False
      'Фильтр по умолчанию настроен на картинки
      If aArg(2) <> vbNullString Then objRegExp.Pattern = aArg(2) Else objRegExp.Pattern = ".+\.(jpg|jpeg|png|emf)$"
    
      For Each FileName In aArg(1)(0).Keys
        If objRegExp.test(FileName) Then
          Set objXML = CreateObject("MSXml2.DOMDocument")
          Set objDocElem = objXML.createElement("Base64Data")
          objDocElem.DataType = "bin.hex"
          objDocElem.nodeTypedValue = aArg(1)(0)(FileName)
          objDocElem.Text = Mid(objDocElem.Text, 41)
          GetValue = objDocElem.nodeTypedValue
          Set objDocElem = Nothing
          Set objXML = Nothing
          Exit For
        End If
      Next
    Case "rtfimg"
      If aArg(0) = 0 Or IsNull(aArg(1)) Or IsEmpty(aArg(1)) Then
        GetValue = vbNullString
      ElseIf LCase(TypeName(aArg(1))) <> "byte()" Then
        GetValue = "{Здесь должно быть изображение: " & aArg(1) & "}"
      Else
        byteStorage = aArg(1)
        GetValue = PictureDataToRTF(byteStorage, aArg(2), aArg(3))
      End If
    Case "fmt"
      GetValue = FormatString(CStr(aArg(1)), ParamList)
    Case "rel"
      GetValue = Trim(aArg(1))
      If ParamList.exists(GetValue) Then
        GetValue = ParamList(GetValue)
        Exit Function
      Else
        Err.Raise 1003, , "Не найдена переменная с именем [" & GetValue & "]"
      End If
    Case "clr"
      For cp = 1 To aArg(0)
        If ParamList.exists(aArg(cp)) Then ParamList.Remove (aArg(cp))
      Next
    Case "sum"
      For cp = 1 To aArg(0) - 1 Step 2
        If ParamList.exists(aArg(cp)) Then aArg(cp + 1) = aArg(cp + 1) + ParamList(aArg(cp))
        ParamList(aArg(cp)) = aArg(cp + 1)
      Next
    Case "inc"
      For cp = 1 To aArg(0)
        If ParamList.exists(aArg(cp)) Then ParamList(aArg(cp)) = ParamList(aArg(cp)) + 1 Else ParamList(aArg(cp)) = 1
      Next
    Case "cts"
      For cp = 1 To aArg(0) - 1 Step 2
        If IsNull(aArg(cp + 1)) Then
          aArg(cp + 1) = 0
        ElseIf IsEmpty(aArg(cp + 1)) Then
          aArg(cp + 1) = 0
        ElseIf aArg(cp + 1) = vbNullString Then
          aArg(cp + 1) = 0
        Else
          aArg(cp + 1) = 1
        End If
        If ParamList.exists(aArg(cp)) Then aArg(cp + 1) = aArg(cp + 1) + ParamList(aArg(cp))
        ParamList(aArg(cp)) = aArg(cp + 1)
      Next
      'дальше описываем остальные функции
    Case "calc", "set"
      GetValue = vbNullString & aArg(1)
      ParamList(GetValue) = aArg(2)
    Case Else
      On Error GoTo TryAsNativeFunction
  
      GetValue = Application.Run(sFnc, ParamList, aArg)
      If Err.Number <> 0 Then
        sErrorMsg = "Ошибка в формуле {" & Formula & "} в позиции " & nvSP & "." & vbCrLf & "[" & Err.Number & "]" & Err.Description
        On Error GoTo 0
        GoTo ResumeOnError
      End If
      Exit Function
TryAsNativeFunction:
      If Err.Number = 2517 Or Err.Number = 450 Or Err.Number = 430 Or Err.Number = 13 Then Resume TryAsNativeFunctionEntry Else GoTo OnError
TryAsNativeFunctionEntry:
      Dim sFunctionCall
  
      sFunctionCall = vbNullString
      For cp = 1 To aArg(0)
        If cp > 1 Then sFunctionCall = sFunctionCall & ", "
        sFunctionCall = sFunctionCall & ToSQL(aArg(cp))
      Next
      sFunctionCall = sFnc & "(" & sFunctionCall & ")"
      On Error GoTo OnError
      GetValue = Application.Eval(sFunctionCall)
    End Select
  Else
    If sFnc = "#" Then
      'Вернет номер строки самого верхнего набора данных
      GetValue = ParamList(ParamList("@SYS_CurrentRecordSet")(0) & ".rownum")
    ElseIf sFnc = "true" Then
      GetValue = True
    ElseIf sFnc = "false" Then
      GetValue = False
    ElseIf sFnc = "null" Then
      GetValue = Null
    ElseIf ParamList.exists(sFnc) Then           'Иначе считаем переменной и ищем в списке
      GetValue = ParamList(sFnc)
      Exit Function
    ElseIf ParamList.exists(Replace(Replace(sFnc, "[", vbNullString), "]", vbNullString)) Then 'Иначе считаем переменной и ищем в списке
      GetValue = ParamList(Replace(Replace(sFnc, "[", vbNullString), "]", vbNullString))
      Exit Function
    Else
      On Error GoTo ParamNotFound
      GetValue = Application.Eval(sFnc)
      'ParamList(sFnc) = GetValue
      Exit Function
ParamNotFound:
      Resume ResumeParamNotFound
ResumeParamNotFound:
      On Error GoTo OnError
      Err.Raise 1002, , "Не удалось получить значение параметра '" & sFnc & "'"
    End If
  End If
  
End If

Exit Function

OnError:
If Err.Number = 1001 Then
  sErrorMsg = Err.Description
Else
  sErrorMsg = "Ошибка в формуле {" & Formula & "} в позиции " & nvSP & "." & vbCrLf & "[" & Err.Number & "]" & Err.Description & vbCrLf & vbCrLf & DumpContext(ParamList)
End If
On Error GoTo 0
Err.Clear
Resume ResumeOnError
ResumeOnError:
Err.Raise 1001, , sErrorMsg

End Function

Public Function GetTemplate(ByVal idReport As Long) As Variant
  'Внутренняя функция для формирования отчета. Извлекает из хранилища шаблон по его коду.
  'Так же производится проверка, если исходный файл существует и его дата изменения больше чем у сохраненного шаблона, то шаблон будет обновлен.
  '#param idReport: Код шаблона

  Dim fso As Object, objF As Object
  Dim tRep As Recordset
  Dim sPathOrig As String, sExtension As String
 
  Set fso = CreateObject("scripting.FileSystemObject")
 
  Set tRep = CurrentDb.OpenRecordset(cReportTable, dbOpenDynaset)
  tRep.FindFirst "id = " & idReport
  
  If tRep.NoMatch Then Err.Raise 1000, , "Не найден шаблон с кодом " & idReport
 
  sPathOrig = tRep("sOrignTemplate")
  If Left(sPathOrig, 2) = ".\" Then sPathOrig = GetPath(CurrentDb.Name) & Mid(sPathOrig, 3)
  sExtension = LCase(fso.GetExtensionName(sPathOrig))
  
  If fso.FileExists(sPathOrig) Then
    Set objF = fso.GetFile(sPathOrig)
    'Debug.Print sPathOrig
    If Nz(tRep.Fields("dEditTemplate").Value, Now) <> objF.DateLastModified Then
      tRep.Edit
      Select Case LCase(sExtension)
      Case "rtf": tRep.Fields("clTemplate").Value = PrepareRTF(sPathOrig)
      End Select
          
      tRep.Fields("dEditTemplate").Value = objF.DateLastModified
      tRep.Update
    End If
    Set objF = Nothing
  End If
  GetTemplate = Array(sExtension, tRep("clTemplate") & vbNullString)
  tRep.Close
  Set tRep = Nothing
 
End Function

'@EntryPoint
Function fncConvertTxtToRTF(ByRef Text As String) As String
  'Внутренняя функция для формирования отчета. Конвертирует текст в корректный RTF блок. Если начинается с `{\*\shppict`, то оставляет как есть
  '#param Text: Текст

  Dim i As Long, ch As String
  If LCase(Left(Text, 11)) = "{\*\shppict" Then
    fncConvertTxtToRTF = Text                    'Не меняем
  Else
    fncConvertTxtToRTF = " "
    For i = 1 To Len(Text)
      ch = Mid(Text, i, 1)
      Select Case Asc(ch)
      Case 123, 125, 92                          '{}\
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

Private Function LongFromByteArray(ByRef Arr, ByVal nOffset)
  'Считывает 32 битное целое из байтового массива
  '#param Arr - массив byte()
  '#param nOffset - смещение

  LongFromByteArray = Arr(nOffset) + Arr(nOffset + 1) * 256 + Arr(nOffset + 2) * 256 * 256 + Arr(nOffset + 3) * 256 * 256 * 256
End Function

Private Function WordFromByteArray(ByRef Arr, ByVal nOffset)
  'Считывает 16 битное целое из байтового массива
  '#param Arr - массив byte()
  '#param nOffset - смещение
  WordFromByteArray = Arr(nOffset) + Arr(nOffset + 1) * 256
End Function

Private Function ByteArrayToHex(ByRef Arr)
  'Преобразует массив байт в последовательность HEX
  '#param Arr - массив byte()
  Dim objDocElem
  Set objDocElem = CreateObject("MSXml2.DOMDocument").createElement("Base64Data")
  objDocElem.DataType = "bin.hex"
  objDocElem.nodeTypedValue = Arr
  ByteArrayToHex = objDocElem.Text
  Set objDocElem = Nothing
End Function

Public Function PictureDataToRTF(ByVal PictureData, ByVal nWidth As Variant, ByVal nHeight As Variant)
  'Внутренняя функция для формирования отчета. Конвертирует изображение в корректный RTF блок.
  '#param PictureData: Байтовый массив с изображением
  '#param nWidth: Целевая ширина картинки
  '#param nHeight: Целевая высота картинки
  '#raises 1010 - Исключение возникает в случае неподходящего формата изображения
   
  If IsNull(PictureData) Then
    PictureDataToRTF = vbNullString
  Else
    PictureDataToRTF = "{\*\shppict{\pict"
    
    If Not IsNull(nWidth) Then PictureDataToRTF = PictureDataToRTF & "\picwgoal" & Int(nWidth * 56.6929133858)
    If Not IsNull(nHeight) Then PictureDataToRTF = PictureDataToRTF & "\pichgoal" & Int(nHeight * 56.6929133858)
    
    
    If PictureData(0) = &H15 And PictureData(1) = &H1C Then
      Dim nProgID, offsProgID, sProgID
      offsProgID = WordFromByteArray(PictureData, 14)
      sProgID = vbNullString
      For nProgID = WordFromByteArray(PictureData, 10) To 2 Step -1
        sProgID = sProgID & Chr(PictureData(offsProgID))
        offsProgID = offsProgID + 1
      Next
      If LCase(sProgID) = "paint.picture" Then
        'Это работает с WordPad но не работает с WinWord
        Dim nOffset
        nOffset = WordFromByteArray(PictureData, 2) + 31
        
        If PictureData(nOffset) = &H42 And PictureData(nOffset + 1) = &H4D Then
        
          nOffset = nOffset + 14
          Dim bitMapSize As Long, picw As Long, pich As Long, bits
          bitMapSize = LongFromByteArray(PictureData, nOffset)
          picw = LongFromByteArray(PictureData, nOffset + 4)
          pich = LongFromByteArray(PictureData, nOffset + 8)
          bits = WordFromByteArray(PictureData, nOffset + 14)
          
          
          bitMapSize = bitMapSize + (((picw * bits + 31) And (Not 31)) \ 8) * pich
          
          PictureDataToRTF = PictureDataToRTF & "\dibitmap0 "
          Dim HexData As String
          HexData = Replace(Replace(Replace(ByteArrayToHex(PictureData), " ", vbNullString), vbCr, vbNullString), vbLf, vbNullString)
          HexData = Mid(HexData, (nOffset) * 2 + 1, (bitMapSize) * 2)
                  
          For nOffset = 1 To Len(HexData) Step 256
            PictureDataToRTF = PictureDataToRTF & Mid(HexData, nOffset, 256) & vbCrLf
          Next
          PictureDataToRTF = PictureDataToRTF & "}}"
          Exit Function
          
        End If
      End If
      
      Err.Raise 1010, , "Не поддерживаемый тип OLE данных"
    End If
    
    
    Select Case GetTypeContent(PictureData)
    Case "jpg": PictureDataToRTF = PictureDataToRTF & "\jpegblip" & vbCrLf
    Case "png": PictureDataToRTF = PictureDataToRTF & "\pngblip" & vbCrLf
    Case "emf": PictureDataToRTF = PictureDataToRTF & "\emfblip" & vbCrLf
    Case "wmf": PictureDataToRTF = PictureDataToRTF & "\wmetafile7" & vbCrLf
    Case Else: Err.Raise 1010, , "Не удалось определить формат изображения"
    End Select
    
    PictureDataToRTF = PictureDataToRTF & ByteArrayToHex(PictureData)
  End If
  PictureDataToRTF = PictureDataToRTF & ByteArrayToHex(PictureData) & "}}"

End Function

Public Function MakeReport(ByRef ts As String, Optional ByRef OutStream As Variant = Nothing, Optional ByRef ParamList As Variant = Nothing) As String
  'Внутренняя функция для формирования отчета. Непосредственное формирование отчета по скомпилированному шаблону
  '#param ts: Шаблон
  '#param OutStream: Выходной поток куда будет записан сформированный документ
  '#param ParamList: Контекст

  'Dim ts As String
  Dim PC As Long, iCnt As Long, iCnt2 As Long, iTmp As Long
  Dim dic
  Dim sSQL As String
  Dim sErrorMsg As String

  Dim aRecordSet As Variant, bisCustomRS As Boolean
 
  Dim l, sKey

 
  Dim aRetCall(128) As Long
  Dim iRC As Long
  Dim sfncConvert As String
  Dim sValue As Variant, sName
 
  If ParamList Is Nothing Then
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
  Else
    Set dic = ParamList
  End If
 
  iRC = 0
 
  If dic("extension") <> vbNullString Then
    sfncConvert = "fncConvertTxtTo" & dic("extension")
  End If
 
  On Error Resume Next
  sValue = Application.Run(sfncConvert, vbNullString)
  If Err.Number <> 0 Then sfncConvert = vbNullString
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
      If Err.Number <> 0 Then
        On Error GoTo 0
        Dim StartPos As Long, sFormula As Variant, sFormat As String
        StartPos = 1
        sFormula = sValue
        sValue = GetExpression(sFormula, dic, StartPos)
        StartPos = StartPos + 1
        If StartPos < Len(sFormula) Then
          sFormat = GetExpression(sFormula, dic, StartPos)
          sValue = Format(sValue, sFormat)
        End If
      Else
        On Error GoTo 0
      End If
   
      sValue = Nz(sValue, vbNullString)
      If sfncConvert <> vbNullString Then sValue = Application.Run(sfncConvert, sValue)
      If Not OutStream Is Nothing Then OutStream.Write sValue Else MakeReport = MakeReport & sValue
      PC = PC + iCnt + 7
    Case "CALC"
      '<DOC>
      '`CALC N[3] V[N]`
      'Выполняет выражение, но значение игнорируется
      '- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
      '- `V` - Выражение длиной заданной в N

      iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
      GetExpression Mid(ts, PC + 7, iCnt), dic, 1
      PC = PC + iCnt + 7
    Case "GOTO"
      '<DOC>
      '`GOTO J[8]`
      'Безусловный прыжок по заранее известному адресу
      '- `J` - Новый адрес
      PC = CLng("&h" & Mid(ts, PC + 4, 8))
    
    Case "JUMP"
      '<DOC>
      '`JUMP N[3] A[N]`
      'Выполняет выражение, значение используется как новый адрес для выполнения
      '- `N` - длина выражения, записанная в виде 3 чисел в шестнадцетиричном виде
      '- `A` - Выражение длиной заданной в N, в котором хранится новый адрес
      iCnt = CInt("&h" & Mid(ts, PC + 4, 3))
      PC = GetExpression(Mid(ts, PC + 7, iCnt), dic, 1)
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
      PC = GetExpression(Mid(ts, PC + 7, iCnt), dic, 1)
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
      If Not ToBool(GetExpression(Mid(ts, PC + 7, iCnt), dic, 1)) Then
        PC = CLng("&h" & Mid(ts, PC + 7 + iCnt, 8))
      Else
        PC = PC + 15 + iCnt
      End If
      '<DOC>
      '`ENDT`
      'Метка конца шаблона. При встрече с данной меткой обработка прервывается
    Case "ENDT"
      Exit Do
      '<DOC>
      '`NOOP`
      'Пустой блок
    Case "NOOP"
      PC = PC + 4
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
      sName = Trim(GetExpression(Mid(ts, PC + 7, iCnt), dic, 1))
    
      iCnt2 = CLng("&h" & Mid(ts, PC + iCnt + 7, 4))
    
      sValue = Mid(ts, PC + iCnt + 11, iCnt2)
      iTmp = 1
      sSQL = LCase(GetPartFromFormula(sValue, iTmp))
      If sSQL = "using" Then
        bisCustomRS = True
        sSQL = LCase(GetPartFromFormula(sValue, iTmp))
        aRecordSet = Array(sName, Nothing, dic("@SYS_CurrentRecordSet"), True)
      
        Set aRecordSet(1) = CreateObject("Scripting.Dictionary")
        aRecordSet(1).CompareMode = 1
        aRecordSet(1)("name") = sSQL
        aRecordSet(1)("alias") = sName
  
        Application.Run sSQL & "_new", aRecordSet(1), Mid(sValue, iTmp), dic
        If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
      Else
        sSQL = Trim(GetExpression(sValue, dic, 1)) 'получаем текст из шаблона
        sSQL = FormatString(sSQL, dic)           'подставляем переменные
      
        'Новый набор данных (Имя набора, RecordSet, Родительский набор данных)
        aRecordSet = Array(sName, Empty, dic("@SYS_CurrentRecordSet"), False)
      
        On Error Resume Next
        If Left(sSQL, 1) = "@" Then
          Dim CurrentForm, FormName
          CurrentForm = Empty
          For Each FormName In Split(Trim(Mid(sSQL, 2)), ".")
            If IsEmpty(CurrentForm) Then
              Set CurrentForm = Forms(FormName)
            Else
              Set CurrentForm = CurrentForm.Form.Controls(FormName)
            End If
            If Err.Number <> 0 Then
              sErrorMsg = "Ошибка при открытии курсора {" & sName & "}." & vbCrLf & Err.Description
              Err.Clear
              On Error GoTo 0
              sErrorMsg = sErrorMsg & vbCrLf & vbCrLf & sSQL & vbCrLf & vbCrLf & DumpContext(ParamList)
              MsgBox sErrorMsg
              Err.Raise 1001, , sErrorMsg
            End If
          Next
          Set aRecordSet(1) = CurrentForm.Form.Recordset.Clone
        Else
          sValue = LCase(GetPartFromFormula(sSQL, 1))
          If sValue = "select" Or sValue = "transform" Or sValue = "" Then
TryAsQuery:
            Set aRecordSet(1) = CurrentDb.OpenRecordset(sSQL) 'Рукописный SQL
          Else
            'Открываем запрос как источник записей. Заполним параметры по имени.
            Dim db As Object, qd As Object, par As Variant
            Set db = CurrentDb
            Set qd = db.QueryDefs(sSQL)
            If Err Then
              Err.Clear
              GoTo TryAsQuery
            End If
            For Each par In qd.Parameters
                par.Value = GetExpression(par.Name, dic)
            Next
            Set aRecordSet(1) = qd.OpenRecordset()
            Set qd = Nothing
            Set db = Nothing
          End If
        End If
      
        If Err.Number <> 0 Then                            'Формируем текст ошибки
          On Error GoTo 0
          sErrorMsg = "Ошибка при открытии курсора {" & sName & "}." & vbCrLf & Err.Description
          Err.Clear
          sErrorMsg = sErrorMsg & vbCrLf & vbCrLf & sSQL & vbCrLf & vbCrLf & DumpContext(ParamList)
          MsgBox sErrorMsg
          Err.Raise 1001, , sErrorMsg
        End If
        On Error GoTo 0
        bisCustomRS = False
      End If
    
    
      Dim bEOF
      If bisCustomRS Then
        bEOF = Application.Run(aRecordSet(1)("name") & "_eof", aRecordSet(1), dic)
        If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
      Else
        bEOF = aRecordSet(1).EOF
      End If
      
      If bEOF Then
        If bisCustomRS Then
          Application.Run aRecordSet(1)("name") & "_close", aRecordSet(1), dic
          If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
        Else
          aRecordSet(1).Close
        End If
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
        Dim objRecordSet, bIsCustom
        Set objRecordSet = aRecordSet(1)
        bIsCustom = aRecordSet(3)
        Set aRecordSet(1) = Nothing
        aRecordSet = aRecordSet(2)
        dic.Remove "@SYS_CurrentRecordSet"
        dic("@SYS_CurrentRecordSet") = aRecordSet
        If bIsCustom Then Application.Run objRecordSet("name") & "_close", aRecordSet(1), dic Else objRecordSet.Close
        
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
      If Err.Number <> 0 Then
        On Error GoTo 0
        sValue = GetExpression(sValue, dic, 1)
      Else
        On Error GoTo 0
      End If
    
      If Not FetchRow(dic, sValue) Then
        'Очиcтим все переменные с указанным префиксом
        sValue = UCase(sValue & ".")
        l = Len(sValue)
        For Each sKey In dic.Keys()
          If UCase(Mid(sKey, 1, l)) = sValue Then
            If UCase(sKey) <> sValue & "EOF" Then dic(sKey) = Empty
          End If
        Next
      
      End If
    Case Else
      Err.Raise 1001, , "Шаблон поломался :(" & vbCrLf & Err.Description
    End Select
  Loop

End Function

Public Function FetchRow(ByRef pDic, Optional ByVal pCursorName As String = vbNullString)
  'Внутренняя функция для формирования отчета. Извлекает из курсора очередную строку и обновляет значения в контексте.
  '#param pDic: Текущий контекст
  '#param pCursorName: Имя курсора. Если не задан то считыается текущий курсор


  Dim vRecordSet, vCursorName, vFiles, tmpdic, fld
  
  vRecordSet = pDic("@SYS_CurrentRecordSet")
  If pCursorName = vbNullString Then
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
  
  If LCase(TypeName(vRecordSet)) = "dictionary" Then
    FetchRow = Application.Run(vRecordSet("name") & "_fetch", vRecordSet, pDic)
  Else
    If Not vRecordSet.EOF Then
      For Each fld In vRecordSet.Fields
        If IsObject(fld.Value) Then
          If LCase(TypeName(fld.Value)) = "recordset2" Then
           
            Set tmpdic = CreateObject("Scripting.Dictionary")
            tmpdic.CompareMode = 1
            
            Set vFiles = fld.Value
            Do While Not vFiles.EOF
              tmpdic(vFiles.Fields("FileName").Value) = vFiles.Fields("FileData").Value
              vFiles.MoveNext
            Loop
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
  End If
  Set vRecordSet = Nothing
End Function

Public Function GetPath(ByVal FullPath As String) As String
  'Возвращает имя директории файла
  '#param FullPath: Полное имя
  Dim lngCurrPos, lngLastPos As Long
  Do
    lngLastPos = lngCurrPos
    lngCurrPos = InStr(lngLastPos + 1, FullPath, "\")
  Loop Until lngCurrPos = 0
  If lngLastPos <> 0 Then GetPath = Left(FullPath, lngLastPos)
End Function

Public Function GetFile(ByVal FullPath As String) As String
  'Возвращает имя файла
  '#param FullPath: Полное имя
  Dim lngCurrPos, lngLastPos As Long
  Do
    lngLastPos = lngCurrPos
    lngCurrPos = InStr(lngLastPos + 1, FullPath, "\")
  Loop Until lngCurrPos = 0
  If lngLastPos <> 0 Then GetFile = Right$(FullPath, Len(FullPath) - lngLastPos)
End Function

Public Function GetExt(ByVal FullPath As String) As String
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
  '#param tpData: Массив байт картинки

  If IsNull(tpData) Then
    GetTypeContent = vbNullString
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
      Else
        GetTypeContent = "не распознан"
      End If
    Else
      GetTypeContent = "не распознан"
    End If
  Else
    GetTypeContent = "не распознан"
  End If

End Function

Function FormatStringA(ByRef Text As String, ParamArray apArgs() As Variant) As String
  'Аналогично FormatString, только принимает на вход неопределенное число параметров, далее переданные параметры будут доступны под именами p1, p2, p3 ....
  Dim vArgs
  vArgs = apArgs
  FormatStringA = FormatString(Text, vArgs)
End Function

Function FormatString(ByRef Text As String, Optional ByRef ParamList As Variant = Nothing) As String
  'Производит замену в тексте подстановочных сиволов на значения заданные в словаре.
  'Подстановочные символы обрамляются символом `%`. Если нужно вывести символ как есть то его необходимо удвоить `%%`.
  'Если первый символ равен @ далее следует имя фильтра и имя поля к которому нужно применить фильтр
  'Если первый символ равен $ далее следует выражение, значение которого нужно вставить в строку в виде SQL литерала
  'В остальных случаях значение между % рассматривается как выражение, значение которого нужно вставить как есть. Дополнительно можно указать формат для предварительной обработки.
  '#param ParamList: Словарь с значениями подстановок

  Dim p As Long, pn As Long, smid As String
  Dim sFilter, i As Long, sField, sOperator, vValue, sAny, vParams
 
 
  Set vParams = Nothing
  If Not IsObject(ParamList) Then
    If IsArray(ParamList) Then
      Set vParams = CreateObject("Scripting.Dictionary")
      vParams.CompareMode = 1
      For i = 0 To UBound(ParamList)
        vParams.Add "p" & (i + 1), ParamList(i)
      Next
    End If
  Else
    Set vParams = ParamList
  End If
  If vParams Is Nothing Then
    Set vParams = CreateObject("Scripting.Dictionary")
    vParams.CompareMode = 1
  End If
 
  FormatString = Text
  p = 1
 
  On Error GoTo ErrorMet
  Do While True
    Do While True
      p = InStr(p, FormatString, "%")
      If p = 0 Then Exit Function
      If Mid(FormatString, p + 1, 1) <> "%" Then Exit Do
      FormatString = Left(FormatString, p) & Mid(FormatString, p + 2)
      p = p + 1
    Loop
    pn = p + 1
    Do While True
      pn = InStr(pn, FormatString, "%")
      If pn = 0 Then Exit Do
      If Mid(FormatString, pn + 1, 1) <> "%" Then Exit Do
      FormatString = Left(FormatString, p) & Mid(FormatString, p + 2)
      p = pn
    Loop
  
    If pn < 1 Then
      Exit Function
    ElseIf pn - p = 1 Then
      FormatString = Left(FormatString, p) & Mid(FormatString, pn + 1)
      p = p + 1
    Else
      smid = Trim(Mid(FormatString, p + 1, pn - p - 1))
   
      If Left(smid, 1) = "@" Then
        'Специальный обработчик фильтров
        i = 2
        sFilter = GetPartFromFormula(smid, i)
     
        Do While Mid(smid, i, 1) = " ": i = i + 1: Loop
        If Mid(smid, i, 1) = "," Or Mid(smid, i, 1) = ";" Then
          i = i + 1
          sField = GetPartFromFormula(smid, i)
          If sField <> vbNullString And sFilter <> vbNullString And vParams.exists(sFilter & ".oper") And vParams.exists(sFilter) Then
            sOperator = UCase(Trim(vParams(sFilter & ".oper")))
            vValue = vParams(sFilter)
            If Not IsArray(vValue) Then vValue = Array(vValue)
            If UBound(vValue) < 1 Then
              Select Case sOperator
              Case "BTW", "BTWWR": sOperator = "GE"
              Case "BTWWL", "BTWWB": sOperator = "GR"
              End Select
            End If
            smid = vbNullString
            If sOperator = "IN" Or sOperator = "NOTIN" Then
              For i = 0 To UBound(vValue)
                If i > 0 Then smid = smid & "," & ToSQL(vValue(i)) Else smid = ToSQL(vValue(i))
              Next
              If smid <> vbNullString Then smid = " AND " & IIf(sOperator = "NOTIN", "NOT ", vbNullString) & sField & " IN (" & smid & ")"
            ElseIf sOperator = "BTW" Or sOperator = "BTWWR" Or sOperator = "BTWWL" Or sOperator = "BTWWB" Then
              smid = " AND " & ToSQL(vValue(0)) & IIf(sOperator = "BTW" Or sOperator = "BTWWR", "<=", "<") & sField & _
                     " AND " & sField & IIf(sOperator = "BTW" Or sOperator = "BTWWL", "<=", "<") & ToSQL(vValue(1))
            Else
              If vParams.exists("%") Then sAny = vParams("%") Else sAny = "%"
              Select Case sOperator
              Case "EQ": smid = " AND " & sField & " = " & ToSQL(vValue(0))
              Case "NE": smid = " AND " & sField & " <> " & ToSQL(vValue(0))
              Case "GR": smid = " AND " & sField & " > " & ToSQL(vValue(0))
              Case "LS": smid = " AND " & sField & " < " & ToSQL(vValue(0))
              Case "GE": smid = " AND " & sField & " >= " & ToSQL(vValue(0))
              Case "LE": smid = " AND " & sField & " <= " & ToSQL(vValue(0))
              Case "CONT": smid = " AND " & sField & " LIKE " & ToSQL(sAny & vValue(0) & sAny)
              Case "NCONT": smid = " AND NOT " & sField & " LIKE " & ToSQL(sAny & vValue(0) & sAny)
              Case "START": smid = " AND " & sField & " LIKE " & ToSQL(vValue(0) & sAny)
              Case Else: smid = Application.Run("Operator_" & sOperator, sField, vValue, vParams) 'пользовательский оператор
              End Select
            End If
          Else
            smid = vbNullString
          End If
          vValue = Empty
        Else
          smid = vbNullString
        End If
      ElseIf Left(smid, 1) = "$" Then
        i = 2
        smid = ToSQL(GetExpression(smid, vParams, i))
      Else
        i = 1
        vValue = GetExpression(smid, vParams, i)
        Do While Mid(smid, i, 1) = " ": i = i + 1: Loop
        If Mid(smid, i, 1) = "," Or Mid(smid, i, 1) = ";" Then
          smid = Format(vValue, GetExpression(smid, vParams, i + 1)) 'Оставшаяся часть используется как формат
        Else
          smid = Format(vValue)
        End If
      End If
   
      If False Then
ErrorMet:                                 smid = "{Error: " & Err.Description & "}"
        Err.Clear
      End If
   
      FormatString = Left(FormatString, p - 1) & smid & Mid(FormatString, pn + 1)
      p = p + Len(smid)

    End If
  Loop
End Function

Function InSet(ByRef spKey As Variant, ParamArray apArgs() As Variant) As Boolean
  'Если первый параметр равен одному из последующих то возвращает Истину. В данной реализации Null = Null возвращает так же Истину
  '#param spKey: Проверяемое значение
  '#param apArgs: Одно или несколько тестовых значений. Если значение является массивом, то проверяется каждый элемент массива. Массивы могут быть вложенными
  Dim v As Variant
  v = apArgs
  InSet = InSetInner(spKey, v)
End Function

Function InSetInner(ByRef spKey As Variant, ByRef apArgs As Variant) As Boolean
  Dim bvIsNull As Boolean
  Dim i As Long
 
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

Public Function ToSQL(ByRef pValue As Variant)
  'На основе типа данных преобразует значение в SQL литерал
  '#param pValue - Значение для преобразования
  Select Case VarType(pValue)
  Case vbString
    ToSQL = "'" & Replace(pValue & "", "'", "''") & "'"
  Case vbDate
    If pValue = CLng(pValue) Then
      ToSQL = "#" & Format(pValue, "mm\/dd\/yyyy") & "#"
    ElseIf pValue < 1 Then
      ToSQL = "#" & Format(pValue, "hh:nn:ss") & "#"
    Else
      ToSQL = "#" & Format(pValue, "mm\/dd\/yyyy hh:nn:ss") & "#"
    End If
  Case vbEmpty, vbNull
    ToSQL = "NULL"
  Case vbBoolean
    If pValue Then ToSQL = "true" Else ToSQL = "false"
  Case vbInteger, vbLong, 20
    ToSQL = pValue & vbNullString
  Case vbSingle, vbDouble, vbCurrency, vbDecimal
    ToSQL = Replace(pValue & vbNullString, ",", ".")
    'vbByte ?? char
  Case Else
    If IsArray(pValue) Then
      Dim vElement
      ToSQL = vbNullString
      For Each vElement In pValue
        If Len(ToSQL) = 0 Then
          ToSQL = ToSQL(vElement)
        Else
          ToSQL = ToSQL & ", " & ToSQL(vElement)
        End If
      Next
      If ToSQL = vbNullString Then ToSQL = "NULL"
    Else
      Err.Raise 1001, , "Unsupported type of SQL value!"
    End If
  End Select
End Function

Public Function SelectOneValue(ByRef sql As String) As Variant
  'Выполняет запрос и значение из первой колонки первой строки
  '#param SQL: Текст запроса
 
  Dim rsdao
  Set rsdao = CurrentDb().OpenRecordset(sql)
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

Function longToByte(ByVal l As Long) As String
  Dim tl: tl = l
  longToByte = Chr(tl Mod 256): tl = tl \ 256
  longToByte = longToByte & Chr(tl Mod 256):  tl = tl \ 256
  longToByte = longToByte & Chr(tl Mod 256):  tl = tl \ 256
  longToByte = longToByte & Chr(tl Mod 256)
End Function

Function intToByte(ByVal i As Long) As String
  Dim ti: ti = i
  intToByte = Chr(ti Mod 256):  ti = ti \ 256
  intToByte = intToByte & Chr(ti Mod 256)
End Function

Function block(ByVal fnc As Long, ByRef Data As String) As String
  block = intToByte(fnc) & Data
  block = longToByte((Len(block) \ 2) + 2) & block
End Function

Function Point(ByVal x As Long, ByVal y As Long) As String
  Point = intToByte(x) & intToByte(y)
End Function

Function color(ByVal r As Long, ByVal g As Long, ByVal b As Long) As String
  color = Chr(0) & Chr(b Mod 256) & Chr(g Mod 256) & Chr(r Mod 256)
End Function

Function RectAsPoligon(ByRef objCount As Long, ByVal l As Long, ByVal t As Long, ByVal r As Long, ByVal b As Long) As String
  objCount = objCount + 1
  RectAsPoligon = block(&H324, intToByte(4) & Point(l, b) & Point(l, t) & Point(r, t) & Point(r, b))
End Function

Function CreatePenIndirect(ByRef objCount As Long, ByVal PenStyle As Long, ByVal pPoint As String, ByVal pColor As String) As String
  objCount = objCount + 1
  CreatePenIndirect = block(&H2FA, intToByte(PenStyle) & pPoint & pColor)
End Function

Function SelectObject(ByVal nObject As Long)
  SelectObject = block(&H12D, intToByte(nObject))
End Function

Function CreateBrushIndirect(ByRef objCount As Long, ByVal style As Long, ByVal color As String, ByVal hatch As Long) As String
  objCount = objCount + 1
  CreateBrushIndirect = block(&H2FC, intToByte(style) & color & intToByte(hatch))
End Function

Sub addInArray(ByRef spArray As Variant, ByRef pItem As Variant)
  'Добавляет значение в массив
  '#param spArray: Массив
  '#param pItem: Добавляемый элемент

  If Not IsArray(spArray) Then spArray = Array()
  ReDim Preserve spArray(UBound(spArray) + 1)
  spArray(UBound(spArray)) = pItem
End Sub

Function zebra2wmf(s, xFactor, yFactor, ByRef MaxWidth)
  Dim recs, objCount As Long, i, l, largest, Size
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
  addInArray recs, block(0, Empty)               'EOF
  zebra2wmf = Join(recs, vbNullString)
  largest = 0
  For Each i In recs
    If Len(i) / 2 > largest Then largest = Len(i) / 2
  Next
  Size = Len(zebra2wmf) / 2 + 9
  zebra2wmf = intToByte(1) & _
              intToByte(9) & _
              intToByte(&H100) & _
              intToByte(Size Mod &H10000) & _
              intToByte(Size \ &H10000) & _
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

Public Function Code128(ByRef pParamList As Object, aArg As Variant) As String
  'REPORT_FUNCTION: Code128(Текст;Ширина;Высота)
  'Вставляет в документ картинку с штрихкодом в формате CODE128. Штрих код должен состоять только из букв английского алфавита и цифр. Контрольное число добавляется автоматически в конец.
  '#param Текст: Кодируемый текст
  '#param Ширина: Целевая ширина штрихкода
  '#param Высота: Целевая высота штрихкода
  '#return: RTF блок
  On Error GoTo OnError:
  Dim byteStorage() As Byte, BCWidth
  If aArg(0) > 0 And aArg(1) <> vbNullString Then
    byteStorage = StrConv(zebra2wmf(code128_zebra(aArg(1), 3), 2, 40, BCWidth), vbFromUnicode)
    Code128 = PictureDataToRTF(byteStorage, aArg(2), aArg(3))
  Else
    Code128 = Empty
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
End Function

Function code128_zebra(SourceString, return_type)
  'Внутренняя функция для формирования штрих кода. Формирует штрих код в формате CODE128
  '#param SourceString: Кодируемый текст
  '#param return_type: Тип возвращаемого результата
  ' {*} 0 - Кодирует для вывода специальным шрифтом
  ' {*} 1 - Формат для чтения человеком
  ' {*} 2 - возвращает контрольню сумму
  ' {*} 3 - Возвращает в виде последовательности символов `|` и ` `


  Dim i, dataToFormat, n, currentEncoding, checkDigitValue, stringlen, currentValue, dataToPrint
  
  If IsNull(SourceString) Then Exit Function
  If SourceString = vbNullString Then Exit Function
 
  
  i = 1
  dataToFormat = Trim(SourceString)
  stringlen = Len(dataToFormat)

 

  If return_type = 1 Then
    'Просто форматируем в переданное значение
    i = 1
    code128_zebra = vbNullString
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
      code128_zebra = Chr(203)                   'A
      currentEncoding = "A"
    ElseIf (Len(dataToFormat) > 4 And isNumber(Mid(dataToFormat, 1, 4))) Or n = 202 Then
      code128_zebra = Chr(205)                   'C
      currentEncoding = "C"
    ElseIf n >= 32 And n < 127 Then
      code128_zebra = Chr(204)                   'B
      currentEncoding = "B"
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
   
    If code128_zebra = vbNullString Then
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
        code128_zebra = vbNullString
        Dim zebraArr: zebraArr = Array(vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "|| ||  ||  ", _
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
                                       "|| || |||| ", "|| |||| || ", "|||| || || ", "| | ||||   ", "| |   |||| ", "|   | |||| ", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, _
                                       vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, _
                                       vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "|| ||  ||  ", "| |||| |   ", "| ||||   | ", _
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
  On Error GoTo OnError:

  Dim byteStorage() As Byte, BCWidth
  If aArg(0) > 0 And aArg(1) <> vbNullString Then
  
    byteStorage = StrConv(zebra2wmf(EAN13_zebra(aArg(1), False), 2, 40, BCWidth), vbFromUnicode)
    EAN13 = PictureDataToRTF(byteStorage, aArg(2), aArg(3))
  Else
    EAN13 = Empty
  End If
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
End Function

Public Function EAN13CheckNumber(ByVal Code)
  'Расчитывает контрольную сумму для штрихкода в формате EAN13
  '#param Code: Число для кодировки
  
  Dim sCode, i, CheckSum
  sCode = Code & vbNullString
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
  sCode = Code & vbNullString
  
  If Not isNumber(sCode) Then Exit Function
  
  
  
  If addCheckSum Then sCode = sCode & EAN13CheckNumber(sCode)

  sCode = Right("0000000000000" & sCode, 13)

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
  On Error GoTo 0
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


'SUBBLOCK_BEGIN:NUMBERITERATION
Public Sub NumberIteration_New(ByRef Container, ByVal sParam, ByRef pDic)
  'Численный итератор, используется в scan. `for using numberIteration from НачальноеЗначение to КонечноеЗначение step Шаг`

  On Error GoTo OnError:
  
  Dim i As Long, tmp
  i = 1
  
  tmp = LCase(GetPartFromFormula(sParam, i))
  If tmp <> "from" Then Err.Raise 1012, , "ожидается from получено [" & tmp & "]"
  Container("current") = CDbl(GetExpression(sParam, pDic, i))
  tmp = LCase(GetPartFromFormula(sParam, i))
  If tmp <> "to" Then Err.Raise 1012, , "ожидается to получено [" & tmp & "]"
  Container("target") = CDbl(GetExpression(sParam, pDic, i))
  tmp = LCase(GetPartFromFormula(sParam, i))
  If tmp = "step" Then
    Container("step") = CDbl(GetExpression(sParam, pDic, i))
  ElseIf tmp <> vbNullString Then
    Err.Raise 1012, , "ожидается step получено [" & tmp & "]"
  Else
    Container("step") = 1
  End If
  
  Exit Sub
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Sub

Public Function NumberIteration_EOF(ByRef Container, ByRef pDic)
  '#hide
  'Численный итератор, проверка на окончание цикла

  On Error GoTo OnError:
  
  If Container("step") < 0 Then
    NumberIteration_EOF = Container("current") <= Container("target")
  Else
    NumberIteration_EOF = Container("current") >= Container("target")
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Function

Public Function NumberIteration_Fetch(ByRef Container, ByRef pDic)
  '#hide
  'Численный итератор, выборка очередного значения

  On Error GoTo OnError:
  
  If NumberIteration_EOF(Container, pDic) Then
    NumberIteration_Fetch = False
    pDic(Container("alias") & ".EOF") = True
  Else
    pDic(Container("alias")) = Container("current")
    Container("current") = Container("current") + Container("step")
    NumberIteration_Fetch = True
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Function

'@EntryPoint
Public Sub NumberIteration_Close(ByRef Container, ByRef pDic)
  '#hide
  'Численный итератор, удаление ресурсов

  On Error GoTo OnError:

  
  Exit Sub
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Sub

'SUBBLOCK_END


'SUBBLOCK_BEGIN:SCAN_WHILE
Public Sub While_New(ByRef Container, ByVal sParam, ByRef pDic)
  'Итератор While, используется в scan. `for using while i < 10`

  On Error GoTo OnError:
  Container("exp") = sParam
  Exit Sub
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Sub

Public Function While_EOF(ByRef Container, ByRef pDic)
  '#hide
  'Итератор While, проверка на окончание цикла

  On Error GoTo OnError:
  While_EOF = Not GetExpression(Container("exp"), pDic, 1)
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Function

Public Function While_Fetch(ByRef Container, ByRef pDic)
  '#hide
  'Итератор While, выборка очередного значения
  While_Fetch = True
  On Error GoTo OnError:
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Function

Public Sub While_Close(ByRef Container, ByRef pDic)
  '#hide
  'Итератор While, удаление ресурсов

  On Error GoTo OnError:
  
  Exit Sub
OnError:
  Dim errNumber, errSource, errDescription
  errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
  
End Sub

'SUBBLOCK_END