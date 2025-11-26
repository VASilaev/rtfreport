Option Compare Binary
' Модуль для конструирования критериев условия where
' Зависит от KRNReport, функция toSQL


Function GetTypeName(targetType As Integer) As String
' Вспомогательная функция для получения имени типа по константе
' #param targetType - тип данных

    Select Case targetType
        Case vbString: GetTypeName = "String"
        Case vbDate: GetTypeName = "Date"
        Case vbBoolean: GetTypeName = "Boolean"
        Case vbInteger: GetTypeName = "Integer"
        Case vbLong: GetTypeName = "Long"
        Case vbSingle: GetTypeName = "Single"
        Case vbDouble: GetTypeName = "Double"
        Case vbCurrency: GetTypeName = "Currency"
        Case vbDecimal: GetTypeName = "Decimal"
        Case Else: GetTypeName = "Unknown"
    End Select
End Function

Public Function BetweenParse(ByRef sOperator As String, ByRef Value As Variant)
'Для оператора Between преобразуетс строку в пару значений. Строка разбивается по строке " и "
' #param sOperator - Текущий оператор, может быть изменен при работе функции
' #param value - Строка с значение, будет преобразована в массив крайних значений, тип останется строчный


  Dim i As Long
  sOperator = UCase(sOperator)
  If Left(Trim(Value), 1) = "(" Then
    Value = Mid(LTrim(Value), 2)
    If sOperator = "BTW" Then
      sOperator = "BTWWL"
    ElseIf sOperator = "BTWR" Then
      sOperator = "BTWWB"
    End If
  End If
  
  If Right(Trim(Value), 1) = "(" Then
    Value = RTrim(Value)
    Value = Left(Value, Len(Value) - 1)
    If sOperator = "BTW" Then
      sOperator = "BTWWR"
    ElseIf sOperator = "BTWL" Then
      sOperator = "BTWWB"
    End If
  End If
  
  i = InStr(1, Value, " и ", vbTextCompare)
  If i > 0 Then
    Value = Array(Left(Value, i - 1), Mid(Value, i + 3))
  End If
End Function

Public Function ConvertValueToType(Value As Variant, targetType As Integer) As Variant
'Конвертирует строковое значение в указанный тип
' #param value - Строка или массив строк
' #param targetType - тип значения к которому нужно привести

    If IsArray(Value) Then
        ' Обработка массива: рекурсивно применить преобразование к каждому элементу
        Dim newArray() As Variant
        Dim i As Long
        Dim lb As Long, ub As Long
        lb = LBound(Value)
        ub = UBound(Value)
        ReDim newArray(lb To ub)
        For i = lb To ub
            newArray(i) = ConvertValueToType(Value(i), targetType)
        Next i
        ConvertValueToType = newArray
    Else
        ' Обработка скалярного значения с обработкой ошибок
        On Error Resume Next
        Select Case targetType
            Case vbString
                ConvertValueToType = CStr(Value)
            Case vbDate
                ConvertValueToType = CDate(Value)
            Case vbBoolean
               Select Case UCase(CStr(Value))
                   Case "ДА", "YES", "Y", "+", "1", "TRUE", "ИСТИНА"
                        ConvertValueToType = True
                   Case Else
                        ConvertValueToType = False
               End Select
            Case vbInteger
                ConvertValueToType = CInt(Value)
            Case vbLong, 20
                ConvertValueToType = CLng(Value)
            Case vbSingle
                ConvertValueToType = CSng(Value)
            Case vbDouble
                ConvertValueToType = CDbl(Value)
            Case vbCurrency
                ConvertValueToType = CCur(Value)
            Case vbDecimal
                ConvertValueToType = CDec(Value)
            Case Else
                Err.Raise 5, "ConvertValueToType", "Неподдерживаемый тип данных"
        End Select
        If Err.Number <> 0 Then
            Dim errorMsg As String
            errorMsg = "Не удалось выполнить преобразование " & CStr(Value) & " к типу " & GetTypeName(targetType)
            Err.Clear
            On Error GoTo 0
            Err.Raise vbObjectError + 1, "ConvertValueToType", errorMsg
        End If
        On Error GoTo 0  ' Восстановить обработку ошибок
    End If
End Function

Public Function GetFieldType(tableOrQueryName As String, fieldName As String, Optional mustFind As Boolean = False) As Integer
' Функция для получения типа поля в таблице или запросе Access по имени таблицы/запроса и поля
' Возвращает константу VBA типа (vbString, vbDate и т.д.) или вызывает ошибку, если поле/таблица/запрос не найдены
' #param tableOrQueryName - имя таблицы или запроса
' #param fieldName - имя поля

    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database

    Dim fld As DAO.Field
    
    ' Открываем текущую базу данных
    Set db = CurrentDb()
    
    ' Сначала пытаемся найти поле в таблицах
    Dim tdf As DAO.TableDef
    On Error Resume Next  ' Временно отключаем обработку ошибок для проверки
    Set tdf = db.TableDefs(tableOrQueryName)
    If Err.Number = 0 Then
        Set fld = tdf.Fields(fieldName)
        If Err.Number = 0 Then
            GoTo DetermineType  ' Найдено в таблице, переходим к определению типа
        End If
    End If
    On Error GoTo ErrorHandler  ' Восстанавливаем обработку ошибок
    
    ' Если не найдено в таблицах, ищем в запросах
    Dim qdf As DAO.QueryDef
    On Error Resume Next
    Set qdf = db.QueryDefs(tableOrQueryName)
    If Err.Number = 0 Then
        Set fld = qdf.Fields(fieldName)
        If Err.Number = 0 Then
            GoTo DetermineType  ' Найдено в запросе, переходим к определению типа
        End If
    End If
    On Error GoTo ErrorHandler
    
    GoTo ErrorHandler
    
    ' Если не найдено ни в таблицах, ни в запросах, ошибка
    Err.Raise vbObjectError + 1, "GetFieldType", "Таблица или запрос '" & tableOrQueryName & "' или поле '" & fieldName & "' не найдены."
    
DetermineType:
    ' Определяем тип поля и возвращаем соответствующую константу VBA
    Select Case fld.Type
        Case dbText
            GetFieldType = vbString
        Case dbMemo
            GetFieldType = vbString
        Case dbDate
            GetFieldType = vbDate
        Case dbBoolean
            GetFieldType = vbBoolean
        Case dbByte  ' Byte может быть обработан как Integer
            GetFieldType = vbInteger
        Case dbInteger
            GetFieldType = vbInteger
        Case dbLong
            GetFieldType = vbLong
        Case dbSingle
            GetFieldType = vbSingle
        Case dbDouble
            GetFieldType = vbDouble
        Case dbCurrency
            GetFieldType = vbCurrency
        Case dbDecimal
            GetFieldType = vbDecimal
        Case Else
            Err.Raise vbObjectError + 1, "GetFieldType", "Неподдерживаемый тип поля: " & fld.Type
    End Select
    
    ' Очистка объектов
    Set fld = Nothing
    Set tdf = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ' Очистка объектов в случае ошибки
    Set fld = Nothing
    Set tdf = Nothing
    Set qdf = Nothing
    Set db = Nothing
    If mustFind Then
      Err.Raise vbObjectError + 1, "GetFieldType", "Ошибка: Не удалось определить тип для '" & tableOrQueryName & "." & fieldName & "'"
    Else
      GetFieldType = -1
    End If
End Function



Public Function BuildFilterClause(sField As String, ByVal sOperator As String, vValue As Variant, Optional sAnyChar As String = "%")
  'Создает фильт для одного поля фильтрации
  '#param sField - имя поля к которому применяется фильтрации
  '#param sOperator - имя оператора, если начинается с NOT то обрабатывается по инверсной логике ( and not ...)
  ' {*} - EQ - равно
  ' {*} - NE - не равно
  ' {*} - DIST - реализуется логика is distinct from (отличается от), eсли в поле значение NULL и сравнивается с не NULL, то результат TRUE
  ' {*} - GR - больше
  ' {*} - LS - меньше
  ' {*} - GE - больше равно
  ' {*} - LE - меньше равно
  ' {*} - CONT - содержит
  ' {*} - START - начинается с
  ' {*} - LIKE - соотетствует
  ' {*} - BTW - между включительно
  ' {*} - BTWWR - между, правая граница не включительно
  ' {*} - BTWWL - между, левая граница не включительно
  ' {*} - BTWWB - между, обе границы не включаются
  

  Dim bNot As Boolean
  If sField <> vbNullString And sOperator <> vbNullString Then
    bNot = False
    If Left(sOperator, 3) = "NOT" Then
      bNot = True: sOperator = Trim(Mid(sOperator, 4))
    End If
    sOperator = UCase(sOperator)
    
    If Not IsArray(vValue) Then vValue = Array(vValue)
    If UBound(vValue) < 1 Then
      Select Case sOperator
      Case "BTW", "BTWWR": sOperator = "GE"
      Case "BTWWL", "BTWWB": sOperator = "GR"
      End Select
    End If
    BuildFilterClause = vbNullString
    If sOperator = "IN" Or sOperator = "NOTIN" Then
      For i = 0 To UBound(vValue)
        If i > 0 Then BuildFilterClause = BuildFilterClause & "," & ToSQL(vValue(i)) Else BuildFilterClause = ToSQL(vValue(i))
      Next
      If BuildFilterClause <> vbNullString Then BuildFilterClause = sField & " IN (" & BuildFilterClause & ")"
    ElseIf sOperator = "BTW" Or sOperator = "BTWWR" Or sOperator = "BTWWL" Or sOperator = "BTWWB" Then
      BuildFilterClause = " AND " & ToSQL(vValue(0)) & IIf(sOperator = "BTW" Or sOperator = "BTWWR", "<=", "<") & sField & _
            sField & IIf(sOperator = "BTW" Or sOperator = "BTWWL", "<=", "<") & ToSQL(vValue(1))
    Else
      Select Case sOperator
      Case "EQ":
        If IsNull(vValue(0)) Or IsEmpty(vValue(0)) Then
          BuildFilterClause = sField & " IS NULL"
        ElseIf VarType(vValue(0)) = vbBoolean Then
           BuildFilterClause = sField
           If Not vValue(0) Then bNot = Not bNot
        Else
          BuildFilterClause = sField & " = " & ToSQL(vValue(0))
        End If
      Case "NE":
        If IsNull(vValue(0)) Or IsEmpty(vValue(0)) Then
          BuildFilterClause = sField & " IS NOT NULL"
        ElseIf VarType(vValue(0)) = vbBoolean Then
           BuildFilterClause = sField
           If vValue(0) Then bNot = Not bNot
        Else
          BuildFilterClause = sField & " <> " & ToSQL(vValue(0))
        End If
      Case "DIST":
        If IsNull(vValue(0)) Or IsEmpty(vValue(0)) Then
          BuildFilterClause = sField & " IS NOT NULL"
        Else
          BuildFilterClause = "(" & sField & " IS NULL OR " & Mid(BuildFilterClause(sField, "NE", vValue), 6) & ")"
        End If
      Case "GR": BuildFilterClause = sField & " > " & ToSQL(vValue(0))
      Case "LS": BuildFilterClause = sField & " < " & ToSQL(vValue(0))
      Case "GE": BuildFilterClause = sField & " >= " & ToSQL(vValue(0))
      Case "LE": BuildFilterClause = sField & " <= " & ToSQL(vValue(0))
      Case "CONT":
        If vValue(0) <> "" Then BuildFilterClause = sField & " LIKE " & ToSQL(sAnyChar & vValue(0) & sAnyChar)
      Case "START":
        If vValue(0) <> "" Then BuildFilterClause = sField & " LIKE " & ToSQL(vValue(0) & sAnyChar)
      Case "LIKE":
        If vValue(0) <> "" Then BuildFilterClause = sField & " LIKE " & ToSQL(vValue(0))
      Case Else: BuildFilterClause = Application.Run("Operator_" & Replace(sOperator, " ", ""), sField, sOperator, vValue, sAnyChar)
      End Select
    End If
    
    If BuildFilterClause <> vbNullString Then BuildFilterClause = " AND " & IIf(bNot, "NOT ", vbNullString) & BuildFilterClause
    
  Else
    BuildFilterClause = vbNullString
  End If
End Function