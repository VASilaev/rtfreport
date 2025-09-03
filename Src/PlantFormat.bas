Option Compare Database
Option Explicit

' Модуль специальной обработки текстов, добавляет функции для использования в шаблоне:
' ef - (EvalFormat) может обрабатывать выражения в тексте
' pf - (PlantFormat) применяет форматирование описанное в тексте аргумента
' Версия 1.0. 2025 год
' Больше информации на странице https://github.com/VASilaev/rtfreport


Const SRV_RTF = 0
Const SRV_EXCEL = 1
Const SRV_NONE = 2

Dim sREPlantFormatBegin As Object
Dim sREPlantFormatBody As Object
Dim CurrentExtension As Integer


Public Function PlantFormat_InitPlugin(Optional ByRef pParamList As Object)
'Добавляет функции в генератор отчетов, а также инициализирует тип сервера печати

  If Not pParamList Is Nothing Then
    Dim sExtension
    sExtension = LCase(pParamList("@SYS_Extension"))
    If sExtension = "rtf" Or sExtension = "docx" Then
      KRNReport.AddSpecialFunction pParamList, "fncEvalTextWrapper", "ef"
      KRNReport.AddSpecialFunction pParamList, "fncPlantFormat", "pf"
      CurrentExtension = SRV_RTF
    ElseIf sExtension = "xls" Or sExtension = "xlsx" Or sExtension = "xlsm" Then
      KRNReport.AddSpecialFunction pParamList, "fncEvalTextWrapper", "ef"
      KRNReport.AddSpecialFunction pParamList, "fncPlantFormat", "pf"
      CurrentExtension = SRV_EXCEL
    End If
  Else
    CurrentExtension = SRV_NONE
  End If
End Function


Function fncPlantFormat(ByRef pParamList As Object, aArg As Variant) As String
' Функция которая воспринимает отдельные конструкции как указания на форматирование и принимает данное форматирование в документе
' Служебные символы перекликаются с разметкой MarkDown
'
' Используемые конструкты:
'  $_подстрочный$
'  $^надстрочный$
'  $__подчеркнутый$
'  $*курсив$
'  $**полужирный$
'  $***полужирный курсив$
'  $~зачеркнутый$
'  $#00ff00 цветной текст$ (только в Excel), цвет задается в формате RGB
'
' Если нужно использовать символ $ с следующим за ним одним из символов (_^*~), то знак доллара нужно удвоить.
' Внутри форматируемого текста символы \ и $ - служебные, и для использования по своему естественному написанию должны экранироваться \\ и \$ соответственно.
'   Например: 2$$*2шт=$~5\$$, преобразуется в текст 2$*2шт=5$, где 5 зачеркнуто
' Форматы можно вкладывать друг в друг
'   Например: $$ - $_$__$#FF0000Подстрочный подчеркнутый и красный$$$



  fncPlantFormat = vbNullString
  On Error GoTo OnError:
  If aArg(0) = 0 Then
    Exit Function
  ElseIf IsNull(aArg(1)) Or IsEmpty(aArg(1)) Then
    Exit Function
  Else
    
    Dim Text, sChunk

    Text = aArg(1)
    If VarType(Text) = vbString Then
      If sREPlantFormatBegin Is Nothing Then
        'Начало формулы
        Set sREPlantFormatBegin = KRNReport.GetRegExp("(?:\\\$|\\\\|\$((__|_|~|\*\*\*|\*\*|\*|\^|#[\da-fA-F]{6}) ?([^\\$]*)))")
        'Тело текста
        Set sREPlantFormatBody = KRNReport.GetRegExp("^([^\\$]*)(\\\\|\\\$|\\|\$(__|_|~|\*\*\*|\*\*|\*|\^|#[\da-fA-F]{6}) ?|\$|$)")
      End If
      
      Dim bNeedFormatter, Stack, objMatches, objMatch, aCurrent
      bNeedFormatter = False
      Set Stack = CreateObject("System.Collections.Stack")
      
      Do While Len(Text) > 0
        
        If Stack.Count = 0 Then
          'Ищем начало формулы
          
          
          Set objMatches = sREPlantFormatBegin.Execute(Text)
          If objMatches.Count = 0 Then Exit Do
          Set objMatch = objMatches.Item(0)
          
          If objMatch.FirstIndex > 0 Then
            fncPlantFormat = fncPlantFormat & Left(Text, objMatch.FirstIndex)
          End If
          
          Text = Mid(Text, objMatch.FirstIndex + objMatch.Length + 1)
          
          If objMatch.Value = "\\" Or objMatch.Value = "\$" Then
            'Экранирование специальных символов
            fncPlantFormat = fncPlantFormat & Right(objMatch.Value, 1)
          Else
            'Старт разметки
            Stack.push Array(objMatch.Submatches(1), objMatch.Submatches(2))
          End If
        Else
          Set objMatches = sREPlantFormatBody.Execute(Text)
          If objMatches.Count = 0 Then Exit Do
          Set objMatch = objMatches.Item(0)
          
          Text = Mid(Text, objMatch.FirstIndex + objMatch.Length + 1)
          
          aCurrent = Stack.Pop
          
          'Добавляем текст
          aCurrent(1) = aCurrent(1) & objMatch.Submatches(0)
          
          Select Case objMatch.Submatches(1)
            Case "", "$"
              'Закрываемся
              
              If CurrentExtension = SRV_RTF Then
                Select Case aCurrent(0)
                  Case "_":   sChunk = Chr(3) & "{\sub " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case "__":  sChunk = Chr(3) & "{\ul " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case "^":   sChunk = Chr(3) & "{\super " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case "***": sChunk = Chr(3) & "{\b\i " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case "**":  sChunk = Chr(3) & "{\b " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case "*":   sChunk = Chr(3) & "{\i " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case "~":   sChunk = Chr(3) & "{\strike " & Chr(2) & aCurrent(1) & Chr(3) & "}" & Chr(2)
                  Case Else
                    'для изменения цвета в rtf нужно менять таблицу цветов
                    sChunk = aCurrent(1)
                End Select
              ElseIf CurrentExtension = SRV_EXCEL Then
                bNeedFormatter = True
              
                Select Case aCurrent(0)
                  Case "_":   sChunk = "<font subscript=true>" & aCurrent(1) & "<font />"
                  Case "__":  sChunk = "<font Underline=true>" & aCurrent(1) & "<font />"
                  Case "^":   sChunk = "<font superscript=true>" & aCurrent(1) & "<font />"
                  Case "***": sChunk = "<font bold=true italic=true>" & aCurrent(1) & "<font />"
                  Case "**":  sChunk = "<font bold=true>" & aCurrent(1) & "<font />"
                  Case "*":   sChunk = "<font italic=true>" & aCurrent(1) & "<font />"
                  Case "~":   sChunk = "<font strikethrough=true>" & aCurrent(1) & "<font />"
                  Case Else
                    If Left(aCurrent(0), 1) = "#" Then
                      sChunk = "<font color=" & RGB(CInt("&H" & Mid(aCurrent(0), 2, 2)), CInt("&H" & Mid(aCurrent(0), 4, 2)), CInt("&H" & Mid(aCurrent(0), 6, 2))) & ">" & aCurrent(1) & "<font />"
                    Else
                      sChunk = aCurrent(1)
                    End If
                End Select
              
              Else
                sChunk = aCurrent(1)
              End If
              
              If Stack.Count > 0 Then
                aCurrent = Stack.Pop
                aCurrent(1) = aCurrent(1) & sChunk
                Stack.push aCurrent
              Else
                fncPlantFormat = fncPlantFormat & sChunk
              End If
            Case "\", "\$", "\\"
              'Закидываем к строке экранированные строки
              aCurrent(1) = aCurrent(1) & Right(objMatch.Submatches(1), 1)
              Stack.push aCurrent
            Case Else
              'Пока закидываем обратно
              Stack.push aCurrent
              'Новый открывающий текст
              Stack.push Array(objMatch.Submatches(2), "")
          End Select
        End If
      Loop
      
      'Остаток текста
      fncPlantFormat = fncPlantFormat & Text
      
      If bNeedFormatter Then ExcelAddCellFormatter pParamList, "fncExcel_FontFormatter", 0
    End If
    
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription

End Function


Function fncEvalText(ByRef Text As String, Optional ByRef ParamList As Variant = Nothing) As String
'Извлекает из текста части заключенные в двойные фигурные скобки, содержимое рассматривается как выражение, которое вычисляется и его значение заменяет формулу.
'
'Например текст "2*2={{2*2}}" будет преобразован в "2*2=4"

  Dim i As Long, ch As String, j As Long, errNum As Long
  

  fncEvalText = " "
  i = 1
  Do While i <= Len(Text)
    ch = Mid(Text, i, 1)
    Select Case Asc(ch)
    Case 123  '{
      If Mid(Text, i + 1, 1) = "{" Then
        j = i + 2
        On Error Resume Next
        ch = GetExpression(Text, ParamList, j)
        errNum = Err.Number
        On Error GoTo 0
        If errNum Then
          fncEvalText = fncEvalText & "\" & ch
        Else
          Do While Mid(Text, j, 1) = " "
            j = j + 1
          Loop
          If Mid(Text, j, 2) = "}}" Then
            i = j + 1
            fncEvalText = fncEvalText & Mid(fncEvalText(ch), 2)
          Else
            'Даже если что то и собрали - это не правильная формула
            fncEvalText = fncEvalText & "\" & ch
          End If
        End If
        
      Else
        fncEvalText = fncEvalText & "\" & ch
      End If
    Case 3 'END OF TEXT (ETX)
      j = InStr(i + 1, Text, Chr(2)) 'START OF TEXT (STX)
      If j = 0 Then j = Len(Text) + 1
      fncEvalText = fncEvalText & Mid(Text, i + 1, j - i - 1)
      i = j
    Case Else
      fncEvalText = fncEvalText & ch
    End Select
    i = i + 1
  Loop

End Function


Function fncEvalTextWrapper(ByRef pParamList As Object, aArg As Variant) As String
'Функция обертка над EvalText для вызова из отчета

  fncEvalTextWrapper = vbNullString
  On Error GoTo OnError:
  
  If aArg(0) = 0 Then
    Exit Function
  ElseIf IsNull(aArg(1)) Or IsEmpty(aArg(1)) Then
    Exit Function
  Else
    PreFormatValue aArg(1), pParamList
    fncEvalTextWrapper = fncEvalText(aArg(1) & "", pParamList)
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription

End Function



Private Function ParseAttributes(attribs As String)
' Разбирает атрибуты в словарь значений
    Dim ParamName As String
    Dim i As Long
    Dim currentFormat As Object
    
    i = 1
    Set currentFormat = CreateObject("Scripting.Dictionary")
    currentFormat.CompareMode = 1
    
    
    Do While True
      ParamName = KRNReport.GetPartFromFormula(attribs, i)
      If ParamName = "" Then Exit Do
      If Mid(attribs, i, 1) <> "=" Then Exit Do Else i = i + 1
      currentFormat(ParamName) = KRNReport.GetValue(attribs, StartPos:=i)
    Loop
      
    Set ParseAttributes = currentFormat
End Function


Public Sub fncExcel_FontFormatter(pRange, ParamList, pUserData)
' анализирует текст внутри ячейки если встречается текст заключенный в тег `<font ...>Текст<font />`, то теги вырезаются,
' а к тексту будет применено форматирование указанное в атрибутах открывающего текста
  On Error GoTo OnError

    
    Dim tagStack As Object, cellText As String, tagStart As Long, tagEnd As Long, tagAttribs As String, Operation As Object, cleartext As String

    If pRange.Value = "" Then Exit Sub

    Set tagStack = CreateObject("System.Collections.Stack")
    Set Operation = CreateObject("System.Collections.Stack")
        
    cellText = pRange.Value
    
    cleartext = ""

    Do While True
      tagStart = InStr(1, cellText, "<font ", vbTextCompare)
      If tagStart = 0 Then Exit Do

      tagEnd = InStr(tagStart, cellText, ">", vbTextCompare)
      If tagEnd = 0 Then Exit Do



      tagAttribs = Mid(cellText, tagStart + 6, tagEnd - tagStart - 6)

      If tagStart > 1 Then
        cleartext = cleartext & Mid(cellText, 1, tagStart - 1)
      End If
      
      cellText = Mid(cellText, tagEnd + 1)

      If Trim(tagAttribs) <> "/" Then
        'открывающий тег
        tagStack.push Array(Len(cleartext) + 1, ParseAttributes(tagAttribs))
      Else
        Dim aStart, formatDict
        aStart = tagStack.Pop
        Operation.push Array(aStart(0), Len(cleartext), aStart(1))


      End If
    Loop
    
    pRange.Value = cleartext & cellText
        
    Do While Operation.Count > 0
      aStart = Operation.Pop
      Set formatDict = aStart(2)
            
      With pRange.Characters(Start:=aStart(0), Length:=aStart(1) - aStart(0) + 1).Font
      
          If formatDict.Exists("subscript") Then .subscript = CBool(formatDict("Subscript"))
          If formatDict.Exists("Underline") Then .Underline = CBool(formatDict("Underline"))
          If formatDict.Exists("Superscript") Then .Superscript = CBool(formatDict("Superscript"))
          If formatDict.Exists("Bold") Then .Bold = CBool(formatDict("Bold"))
          If formatDict.Exists("italic") Then .Italic = CBool(formatDict("italic"))
          If formatDict.Exists("Strikethrough") Then .Strikethrough = CBool(formatDict("Strikethrough"))
          If formatDict.Exists("color") Then .color = CLng(formatDict("color"))
      
          ' Обработка стандартных свойств
          If formatDict.Exists("Name") Then .name = formatDict("Name")
          If formatDict.Exists("Size") Then .Size = Val(formatDict("Size"))
          
      End With
    Loop
        
        
  Exit Sub
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription

End Sub