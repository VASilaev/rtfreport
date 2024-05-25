Option Compare Database
Option Explicit

Private Function ResolvePath(ObjFSO As Object, ByVal sBasePath As String, ByVal sFileName As String)
'Вычисляет абсолютный путь от относительного
  Dim svBase As String
  ResolvePath = sFileName
  svBase = sBasePath
  If Left(ResolvePath, 2) = ".\" Then
    ResolvePath = Mid(ResolvePath, 3)
    Do While Left(ResolvePath, 3) = "..\"
      ResolvePath = Mid(ResolvePath, 4)
      svBase = ObjFSO.GetParentFolderName(svBase)
    Loop
    ResolvePath = ObjFSO.buildpath(svBase, ResolvePath)
  End If
End Function

Public Function TestClearError(tContext As Object)
'Удаляет из котекста сообщение об ошибке
  tContext("ErrorLevel") = 0
  tContext("ErrorMessage") = vbNullString
  TestClearError = True
End Function

Public Function ModuleExist(tContext As Object)
'Проверяет наличие модуля в проекте VBA. Имя модуля берется из контекста с именем `sModuleName`

  ModuleExist = False
  Dim objModule, sModuleName
  sModuleName = LCase(tContext("sModuleName"))
  For Each objModule In CurrentProject.AllModules
    If LCase(objModule.Name) = sModuleName Then
        ModuleExist = True
        Exit For
    End If
  Next
End Function

Public Function ImportModuleIfNotExists(tContext As Object)
'Проверяет наличие модуля в проекте VBA. Имя модуля берется как имя файла из контекста с именем `sFile`. Если модуля нет, то импортирует его из файла
  ImportModuleIfNotExists = False
  Dim ObjFSO As Object
  Set ObjFSO = tContext("FileSystemObject")
  tContext("sModuleName") = ObjFSO.GetBaseName(tContext("sFile"))
  If Not ModuleExist(tContext) Then
    Application.LoadFromText acModule, ObjFSO.GetBaseName(tContext("sFile")), ResolvePath(ObjFSO, ObjFSO.GetParentFolderName(tContext("sCurrentFileName")), tContext("sFile"))
    ImportModuleIfNotExists = True
  End If
End Function

Public Function CompileAndSaveAllModules(tContext As Object)
'Компилирует и сохраняет все модули. (не работает если есть ошибки)
  DoCmd.RunCommand acCmdCompileAndSaveAllModules
  CompileAndSaveAllModules = True
End Function

Public Function ApplicationQuit(tContext As Object)
'Компилирует и сохраняет все модули. (не работает если есть ошибки)
  Application.Quit tContext("Save")
  ApplicationQuit = True
End Function

Public Function ImportForm(tContext As Object) As Boolean
'Импортирует форму. Имя формы берется как имя файла из контекста с именем `sFile`
  Dim ObjFSO As Object
  Set ObjFSO = tContext("FileSystemObject")
  Application.LoadFromText acForm, ObjFSO.GetBaseName(tContext("sFile")), ResolvePath(ObjFSO, ObjFSO.GetParentFolderName(tContext("sCurrentFileName")), tContext("sFile"))
  ImportForm = True
End Function

Public Function ExportModule(tContext As Object) As Boolean

  Dim ObjFSO As Object
  Set ObjFSO = tContext("FileSystemObject")
  Application.SaveAsText acModule, ObjFSO.GetBaseName(tContext("sFile")), ResolvePath(ObjFSO, ObjFSO.GetParentFolderName(tContext("sCurrentFileName")), tContext("sFile"))
  ExportModule = True
End Function

Public Function ExportForm(tContext As Object) As Boolean
  Dim ObjFSO As Object
  Set ObjFSO = tContext("FileSystemObject")
  Application.SaveAsText acForm, ObjFSO.GetBaseName(tContext("sFile")), ResolvePath(ObjFSO, ObjFSO.GetParentFolderName(tContext("sCurrentFileName")), tContext("sFile"))
  ExportForm = True
End Function

Public Function ResolvePathFromCurrentFile(tContext As Object) As Boolean
  tContext("sFile") = ResolvePath(tContext("FileSystemObject"), tContext("FileSystemObject").GetParentFolderName(tContext("sCurrentFileName")), tContext("sFile"))
  ResolvePathFromCurrentFile = True
End Function

Public Function ExecuteStatement(tContext As Object) As Boolean
  CurrentProject.Connection.Execute tContext("sStatement")
  ExecuteStatement = True
End Function

Public Sub RunScript(tContext As Object)
'Парсер тестовых файлов

Dim fso
Set fso = CreateObject("scripting.FileSystemObject")

Dim tf, nLine, sFileName, sCurrentFile
If tContext.exists("sCurrentFileName") Then sCurrentFile = tContext("sCurrentFileName") Else sCurrentFile = Empty
sFileName = tContext("sFileName")
If Left(sFileName, 2) = ".\" Then sFileName = GetPath(CurrentDb.Name) & Mid(sFileName, 3)
tContext("sCurrentFileName") = sFileName
tContext("Description") = fso.GetBaseName(sFileName)
Set tf = fso.OpenTextFile(sFileName)
nLine = 0

Dim bOnErrorResume As Boolean
bOnErrorResume = False

Dim nCurrentIndent ', nSkipIndent
Dim i As Long, vValue, sVarName

'nSkipIndent = 0

Do While Not tf.AtEndOfStream
  Dim sLine
  
  sLine = tf.ReadLine
  nCurrentIndent = Len(sLine)
  sLine = LTrim(sLine)
  nCurrentIndent = nCurrentIndent - Len(sLine)
  nLine = nLine + 1
  
mProcessLine:

  If Left(sLine, 1) = "'" Or Trim(sLine) = vbNullString Then
    nLine = nLine 'Ни чего не делаем комментарий или пустая строка
  ElseIf Left(sLine, 1) = "@" Then
    'Выполнить действие
    If bOnErrorResume Then
      On Error Resume Next
    End If
    tContext("sys_Result") = Application.Run(Trim(Mid(sLine, 2)), tContext)
    If bOnErrorResume Then
      If Err Then
        tContext("ErrorLevel") = Err.Number
        tContext("ErrorMessage") = "[" & Err.Number & "] " & Err.Description
        Err.Clear
      End If
      On Error GoTo 0
    End If
  ElseIf Left(sLine, 1) = "#" Then
    'Специальные команды
    sLine = LCase(Trim(Mid(sLine, 2)))
    If sLine = "onerrorresume" Then
      bOnErrorResume = True
    ElseIf sLine = "onerrorstop" Then
      bOnErrorResume = False
    ElseIf Left(sLine, 3) = "if(" Then
      i = 4
      vValue = GetExpression(sLine, tContext, i)
      If Mid(sLine, i, 1) <> ")" Then Err.Raise "Ожидается )"
      sLine = Mid(sLine, i + 1)
      If vValue Then GoTo mProcessLine
    End If
  Else
    i = InStr(sLine, "=")
    If i = 0 Then
      Err.Raise 2000, , "Не удалось распознать действие"
    Else
      sVarName = Trim(Left(sLine, i - 1))
      vValue = LTrim(Mid(sLine, i + 1))
      Select Case Left(vValue, 1)
        Case "&": tContext(sVarName) = CLng(Mid(vValue, 2))
        Case "!": tContext(sVarName) = CDbl(Mid(vValue, 2))
        Case "#": tContext(sVarName) = CDate(Mid(vValue, 2))
        Case "%"
          tContext("nLine") = nLine
          If bOnErrorResume Then
            On Error Resume Next
          End If
          tContext(sVarName) = GetExpression(vValue, tContext, 2)
          If bOnErrorResume Then
            If Err.Number <> 0 Then
              tContext("ErrorLevel") = Err.Number
              tContext("ErrorMessage") = "[" & Err.Number & "] " & Err.Description
              Err.Clear
            End If
            On Error GoTo 0
          End If
          
        Case "$"
          Dim sMet
          i = InStr(2, vValue, "$")
          If i = 0 Then
            Err.Raise 2001, , "Не верно оформлен текстовый литерал"
          Else
            sMet = Left(vValue, i)
            vValue = Mid(vValue, i + 1)
            Do While Not tf.AtEndOfStream
              sLine = LTrim(tf.ReadLine)
              nLine = nLine + 1
              If Trim(sLine) = sMet Then Exit Do
              vValue = vValue & vbCrLf & sLine
            Loop
            If Trim(sLine) = sMet Then tContext(sVarName) = vValue
          End If
        Case Else: tContext(sVarName) = vValue
      End Select
    End If
  End If
Loop

If Not IsEmpty(sCurrentFile) Then tContext("sCurrentFileName") = sCurrentFile Else tContext.Remove ("sCurrentFileName")

End Sub

Public Function AssertEquals(tContext As Object) As Boolean
'Из контекста берет значения `Expected` и `Actual`. Если значения различаются то формирует запись о проваленном тесте. Имя Теста задается в переменной `Description`
    AssertEquals = True
    If tContext("Expected") <> tContext("Actual") Then
        PrintMessage tContext, tContext("Expected"), tContext("Actual")
        AssertEquals = False
    End If
    tContext("CountRuns") = tContext("CountRuns") + 1
End Function

Private Sub PrintFailDescription(tContext As Object)
'Внутренняя функция вывода заголовка проваленного теста
    tContext("CountFailures") = tContext("CountFailures") + 1
    Dim sDescription
    sDescription = tContext("Description")
    If tContext("PrevDescription") <> sDescription Then
        tContext("OutputMessage").Add "--------------------------------------------------------------"
        tContext("OutputMessage").Add "ПРОВАЛ: " & sDescription
        tContext("OutputMessage").Add "--------------------------------------------------------------"
        tContext("PrevDescription") = sDescription
    Else
        tContext("OutputMessage").Add Chr(9) & "----------------------------------------------------------"
    End If
End Sub

Private Sub PrintMessage(tContext As Object, Expected As Variant, But As Variant)
'Выводит сообщение о проваленном тесте
    PrintFailDescription tContext
    tContext("OutputMessage").Add Chr(9) & "   Строка: " & tContext("nLine")
    tContext("OutputMessage").Add Chr(9) & "Ожидалось: " & Expected
    tContext("OutputMessage").Add Chr(9) & " Получено: " & But
End Sub

Public Sub DebugOutput(tContext As Object, sMessage)
'Вывод сообщений в Debug
  Debug.Print sMessage
End Sub

Public Sub FileOutput(tContext As Object, sMessage)
'Вывод сообщений в открытый файл
  tContext("OtputFile").Write sMessage & vbCrLf
End Sub

Public Sub PrintSummary(tContext As Object)
'Вывод сводной инфоромаци о пройденных тестах
  Dim sOutputProvider
  If tContext.exists("sOutputProvider") Then sOutputProvider = tContext("sOutputProvider") Else sOutputProvider = "DebugOutput"
  

  If tContext("CountRuns") > 0 Then Application.Run sOutputProvider, tContext, "Запущено тестов: " & tContext("CountRuns")
  If tContext("CountFailures") > 0 Then Application.Run sOutputProvider, tContext, "ПРОВАЛЕНО: " & tContext("CountFailures")
  If tContext("CountFailures") = 0 And tContext("CountRuns") > 0 Then Application.Run sOutputProvider, tContext, "OK"
  Dim sMessages
  For Each sMessages In tContext("OutputMessage")
    Application.Run sOutputProvider, tContext, sMessages
  Next
End Sub


Public Function InitializeScripterContext(Optional tContext As Object = Nothing)
'BИнициализация контекста
  If tContext Is Nothing Then
    Set tContext = CreateObject("Scripting.Dictionary")
    tContext.CompareMode = 1
  End If

  tContext.Add "FileSystemObject", CreateObject("scripting.FileSystemObject")
    
  Set InitializeScripterContext = tContext
End Function

Public Function InitializeTestContext(Optional tContext As Object = Nothing)
'Инициализация контекста для запуска тестов
  If tContext Is Nothing Then
    Set tContext = CreateObject("Scripting.Dictionary")
    tContext.CompareMode = 1
  End If
  tContext.Add "Description", "Без описания"
  tContext.Add "PrevDescription", vbNullString
  tContext.Add "OutputMessage", New Collection
  tContext.Add "CountFailures ", 0
  tContext.Add "CountRuns", 0
  TestClearError tContext
  Set InitializeTestContext = tContext
End Function

Public Sub OpenOutputFile(tContext As Object)
'Открывает файл `sFile` и его объект помещается в контекст под именем `OtputFile`
  tContext.Add "OtputFile", ResolvePath(tContext("FileSystemObject"), tContext("FileSystemObject").GetParentFolderName(CurrentDb.Name), tContext("sFile"))
End Sub

Public Function RunFileFromEnv()
'Получет из олкруженя значение пременной %SCRIPT_FILE%" и отправляет ее на выполнение
'используется совместно с макросом RunFileFromEnv. Для автоматического выполнения скрипта из командной строки
'
'```
'set SCRIPT_FILE=.\Tests\ExportModules.txt
'Платежи.accdb /nostartup /x RunFileFromEnv
'```
  RunScriptFile CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SCRIPT_FILE%")
  RunFileFromEnv = True
End Function

Public Function RunScriptFile(sFileName)
'Запускает указанный файл на выполнение
  Dim tContext As Object, nExitWithCode As Variant
  Set tContext = InitializeScripterContext()
  tContext("sFileName") = ResolvePath(tContext("FileSystemObject"), GetPath(CurrentDb.Name), sFileName)
  RunScript tContext
  If tContext.exists("nExitWithCode") Then nExitWithCode = tContext("nExitWithCode") Else
  tContext.RemoveAll
  If Not IsEmpty(nExitWithCode) Then Application.Quit nExitWithCode
  RunScriptFile = True
End Function

Public Sub RunAllTest()
'Запускает все тесты из папки tests
  Dim tContext As Object, sFile
  Set tContext = InitializeTestContext(InitializeScripterContext())
  
  tContext.Add "OtputFile", tContext("FileSystemObject").CreateTextFile(GetPath(CurrentDb.Name) & "tests\Result.log", True)
  
  For Each sFile In tContext("FileSystemObject").getfolder(GetPath(CurrentDb.Name) & "\tests").Files
    If LCase(tContext("FileSystemObject").GetExtensionName(sFile.Path)) = "txt" Then
      tContext("sFileName") = sFile.Path
      RunScript tContext
    End If
  Next
    
  tContext("sOutputProvider") = "FileOutput"
  PrintSummary tContext
  tContext.RemoveAll
End Sub