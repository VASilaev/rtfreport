Option Compare Database
Option Explicit


' Модуль генерации отчетов в формате Excel из шаблона
' Версия 1.0. 2024 год
' Больше информации на странице https://github.com/VASilaev/rtfreport


Function ParseTemplateFormula(ByVal sValue)
'Парсим вероятную формулу на элементы
  Dim SlashPos, BreakerPos, Breaker, aFormula, prevItemType, i
  prevItemType = -1 ' Не известно

    'Парсим выражение, возможно это даже не формула. Формула должна содержать закрытые {}, в качестве символа экранирования используем "\"
  aFormula = Array()
  Breaker = "{"
  Do While True
    SlashPos = InStr(sValue, "\")
    BreakerPos = InStr(sValue, Breaker)
    If SlashPos = 0 And BreakerPos = 0 Then Exit Do
    If SlashPos = 0 Then SlashPos = Len(sValue) + 1
    If BreakerPos = 0 Then BreakerPos = Len(sValue) + 1
    
    If SlashPos < BreakerPos Then
      If SlashPos > 1 Then
        If prevItemType <> -1 Then
          aFormula(UBound(aFormula))(1) = aFormula(UBound(aFormula))(1) & Left(sValue, SlashPos - 1)
        Else
          ReDim Preserve aFormula(UBound(aFormula) + 1)
          prevItemType = 0
          aFormula(UBound(aFormula)) = Array(prevItemType, Left(sValue, SlashPos - 1))
        End If
        sValue = Mid(sValue, SlashPos)
      End If
      
      If Left(sValue, 2) = "\\" Or Left(sValue, 2) = "\{" Or Left(sValue, 2) = "\}" Then i = 2 Else i = 1
      If prevItemType <> -1 Then
        aFormula(UBound(aFormula))(1) = aFormula(UBound(aFormula))(1) & Mid(sValue, i, 1)
      Else
        ReDim Preserve aFormula(UBound(aFormula) + 1)
        prevItemType = 0
        aFormula(UBound(aFormula)) = Array(prevItemType, Mid(sValue, i, 1))
      End If
      sValue = Mid(sValue, i + 1)

    ElseIf Breaker = "{" Then
      If BreakerPos > 1 Then
        If prevItemType <> -1 Then
          aFormula(UBound(aFormula))(1) = aFormula(UBound(aFormula))(1) & Left(sValue, BreakerPos - 1)
        Else
          ReDim Preserve aFormula(UBound(aFormula) + 1)
          prevItemType = 0
          aFormula(UBound(aFormula)) = Array(prevItemType, Left(sValue, BreakerPos - 1))
        End If
      End If
      sValue = Mid(sValue, BreakerPos + 1)
      
      prevItemType = 1
      ReDim Preserve aFormula(UBound(aFormula) + 1)
      aFormula(UBound(aFormula)) = Array(prevItemType, "")
      Breaker = "}"
    Else
      aFormula(UBound(aFormula))(1) = aFormula(UBound(aFormula))(1) & Left(sValue, BreakerPos - 1)
      prevItemType = -1
      Breaker = "{"
      sValue = Mid(sValue, BreakerPos + 1)
    End If
  Loop
  
  If Len(sValue) > 0 Then
    If prevItemType = -1 Then
      ReDim Preserve aFormula(UBound(aFormula) + 1)
      aFormula(UBound(aFormula)) = Array(0, sValue)
    ElseIf Breaker = "{" Then
      aFormula(UBound(aFormula))(1) = aFormula(UBound(aFormula))(1) & sValue
    Else 'Формула не закрыта значит это просто текст
      aFormula(UBound(aFormula))(1) = "{" & aFormula(UBound(aFormula))(1) & sValue
      aFormula(UBound(aFormula))(0) = 0
      If UBound(aFormula) > 0 Then
        If aFormula(UBound(aFormula) - 1)(0) = 0 Then
          aFormula(UBound(aFormula) - 1)(1) = aFormula(UBound(aFormula) - 1)(1) & aFormula(UBound(aFormula))(1)
          ReDim Preserve aFormula(UBound(aFormula) - 1)
        End If
      End If
    End If
  End If

  If UBound(aFormula) < 0 Then aFormula = Array(Array(0, ""))

  ParseTemplateFormula = aFormula
End Function

Function ExcelReportGetModel(objSheet)
'Извлекаем с листа модель заполнения
  Dim model, workbookname, rng
  model = Array()
  
  workbookname = objSheet.name
  
  
  For Each rng In objSheet.Names
    'Находим именованные диапазоны заданного формата
    If LCase(Right(rng.name, 7)) = ".record" Then
      ReDim Preserve model(UBound(model) + 1)
      model(UBound(model)) = Array(rng.RefersToRange.Row, rng.RefersToRange.Column, 0, 1, 1, LCase(Mid(rng.name, 1, Len(rng.name) - 7)), rng.Comment)
        
      If IsArray(rng.RefersToRange.Value2) Then
        model(UBound(model))(3) = UBound(rng.RefersToRange.Value2, 1)
        model(UBound(model))(4) = UBound(rng.RefersToRange.Value2, 2)
      End If
    End If
  Next
  
  For Each rng In objSheet.Parent.Names
    'Находим именованные диапазоны заданного формата
    If LCase(Right(rng.name, 7)) = ".record" And Left(rng.RefersTo, Len(workbookname) + 2) = "=" & workbookname & "!" Then
      ReDim Preserve model(UBound(model) + 1)
      
      model(UBound(model)) = Array(rng.RefersToRange.Row, rng.RefersToRange.Column, 0, 1, 1, LCase(Mid(rng.name, 1, Len(rng.name) - 7)), rng.Comment)
        
      If IsArray(rng.RefersToRange.Value2) Then
        model(UBound(model))(3) = UBound(rng.RefersToRange.Value2, 1)
        model(UBound(model))(4) = UBound(rng.RefersToRange.Value2, 2)
      End If
    End If
  Next
  
  Dim i, j
  
  'Рекрдсеты не должны пересекаться
  For i = LBound(model) To UBound(model) - 1
    For j = i + 1 To UBound(model)
      If model(i)(0) < model(j)(0) + model(j)(3) And model(i)(0) + model(i)(3) > model(j)(0) Then
        Err.Raise 2000, , "Набор данных [" & model(i)(5) & "] пересекается с [" & model(j)(5) & "]"
      End If
    Next
  Next
  
  'Собираем формулы
  Dim FindedCell, aFormula, StartAddress
  Set FindedCell = objSheet.Cells.Find("*{*}*")
  Do While Not FindedCell Is Nothing
    If IsEmpty(StartAddress) Then StartAddress = FindedCell.Address Else If StartAddress = FindedCell.Address Then Exit Do
    
    aFormula = ParseTemplateFormula(FindedCell.Text)
      
    'Если не формула - игнорируем
    If Not (UBound(aFormula) = 0 And aFormula(0)(0) = 0) Then
      ReDim Preserve model(UBound(model) + 1)
      model(UBound(model)) = Array(FindedCell.Row, FindedCell.Column, 1, aFormula)
    End If
    Set FindedCell = objSheet.Cells.FindNext(FindedCell)
  Loop
  
  'Сортируем элементы
  Dim swap
  For i = LBound(model) To UBound(model) - 1
    For j = i + 1 To UBound(model)
      swap = False
            
      If model(i)(2) = 0 And model(j)(2) = 1 Then
        If model(j)(0) >= model(i)(0) And model(j)(0) < model(i)(0) + model(i)(3) And (model(j)(1) < model(i)(1) Or model(j)(1) >= model(i)(1) + model(i)(4)) Then
          'Специальный случай, формулы которые попали в строки рекордсета, но находятся вне его диапазона по столбцам должны обработаться до самого рекордсета
          swap = True
        End If
      End If
      
      If Not swap And (model(j)(0) < model(i)(0) Or _
             (model(j)(0) = model(i)(0) And model(j)(1) < model(i)(1)) Or _
             (model(j)(0) = model(i)(0) And model(j)(1) = model(i)(1) And model(j)(2) < model(i)(2))) Then
        'Заполняемые ячейки заполняются сверху вниз, слева направо. В первую очередь обрабатывается рекордсет.
        swap = True
      End If
      
      If swap Then
        swap = model(i)
        model(i) = model(j)
        model(j) = swap
      End If
    Next
  Next
  
  ExcelReportGetModel = model
End Function


Public Function ExcelReportFillSheet(objSheet, aModel, ParamList, Optional ByVal nRowStart = -1, Optional ByVal nRowEnd = -1)
  
  Dim i, j, FormulaValue, FormulaElement, aRecordSet
  If nRowStart = -1 Then nRowStart = LBound(aModel)
  If nRowEnd = -1 Then nRowEnd = UBound(aModel)
  
  For i = nRowStart To nRowEnd
    If aModel(i)(2) = 1 Then
      'Заполнение формул вида {формула}
      FormulaValue = Empty
      
      If ParamList.exists("@SYS_CurrentCell") Then ParamList.Remove "@SYS_CurrentCell"
      ParamList.Add "@SYS_CurrentCell", objSheet.Cells(aModel(i)(0), aModel(i)(1))
      
      For Each FormulaElement In aModel(i)(3)
        If FormulaElement(0) = 0 Then
          FormulaValue = FormulaValue & FormulaElement(1)
        ElseIf FormulaElement(0) = 1 Then
          If IsEmpty(FormulaValue) Then
            FormulaValue = GetExpression(FormulaElement(1), ParamList)
          Else
            FormulaValue = FormulaValue & GetExpression(FormulaElement(1), ParamList)
          End If
        Else
          Err.Raise 2002, , "Что то пошло не так модель сломалась"
        End If
      Next
      objSheet.Cells(aModel(i)(0), aModel(i)(1)).Value = FormulaValue
      
      ParamList.Remove "@SYS_CurrentCell"
      
    ElseIf aModel(i)(2) = 0 Then
      'Обработка наборов данных, вложенные отсекаются на уровне модели
      Dim CurrentOffsetRow, ElementsInDataset, bShift
      
      CurrentOffsetRow = 0
      ElementsInDataset = 0
      For j = i + 1 To UBound(aModel)
        If aModel(j)(0) >= aModel(i)(0) + aModel(i)(3) Then Exit For
        ElementsInDataset = ElementsInDataset + 1
      Next
      
      'OpenRecordsetForReport возвращает EOF
      If Not OpenRecordsetForReport(aModel(i)(5), aModel(i)(6), ParamList, aRecordSet) Then
        Do While FetchRow(ParamList)
          If Not EOFRecordsetForReport(ParamList, aRecordSet) Then
            objSheet.rows((aModel(i)(0) + CurrentOffsetRow) & ":" & (aModel(i)(0) + aModel(i)(3) - 1 + CurrentOffsetRow)).Select
            objSheet.Application.selection.copy
            objSheet.Application.selection.Insert (-4121) 'xlDown
            bShift = True
          Else
            bShift = False
          End If
          
          'Помещаем в контекст текущую строку
          If ParamList.exists("@SYS_CurrentRow") Then ParamList.Remove "@SYS_CurrentRow"
          ParamList.Add "@SYS_CurrentRow", objSheet.Range(objSheet.Cells(aModel(i)(0) + CurrentOffsetRow, aModel(i)(1)), objSheet.Cells(aModel(i)(0) + CurrentOffsetRow + aModel(i)(3) - 1, aModel(i)(1) + aModel(i)(4) - 1))
          If ParamList.exists("@SYS_CurrentCell") Then ParamList.Remove "@SYS_CurrentCell"
                
          'Заполняем поля
          If ElementsInDataset > 0 Then ExcelReportFillSheet objSheet, aModel, ParamList, i + 1, i + ElementsInDataset
                  
          'Сдвигаем всех на высоту блока
          If bShift Then
            For j = i + 1 To UBound(aModel)
              aModel(j)(0) = aModel(j)(0) + aModel(i)(3)
            Next
          End If
          
          CurrentOffsetRow = CurrentOffsetRow + aModel(i)(3)
        Loop
      Else
        'Удаляем шаблон
        objSheet.rows((aModel(i)(0) + CurrentOffsetRow) + ":" + (aModel(i)(0) + CurrentOffsetRow + aModel(i)(3) - 1)).Select
        objSheet.selection.Delete (-4162) 'xlUp
        'Сдвигаем в обратном направлении
        For j = i + 1 To UBound(aModel)
          aModel(j)(0) = aModel(j)(0) - aModel(i)(3)
        Next
      End If
      'Пропускаем элементы из цикла
            
      If ParamList.exists("@SYS_CurrentRow") Then ParamList.Remove "@SYS_CurrentRow"
     
      'Помистим в контекст обработанный диапазон
      If ParamList.exists("@SYS_PrevRecordset") Then ParamList.Remove "@SYS_PrevRecordset"
      If CurrentOffsetRow = 0 Then
        ParamList.Add "@SYS_PrevRecordset", Nothing
      Else
        ParamList.Add "@SYS_PrevRecordset", objSheet.Range(objSheet.Cells(aModel(i)(0), aModel(i)(1)), objSheet.Cells(aModel(i)(0) + CurrentOffsetRow - 1, aModel(i)(1) + aModel(i)(4) - 1))
      End If
      
      CloseRecordsetForReport ParamList, aRecordSet
      
      i = i + ElementsInDataset
    Else
      Err.Raise 2001, , "Что то пошло не так модель сломалась"
    End If
  Next
  
End Function


Private Function MakeReportExcel(Template, ParamList, sOutFile, bPrint)
  Dim WorkBook, Sheet
  
  If ParamList.exists("@SYS_FncList") Then
    ParamList("@SYS_FncList").RemoveAll
    ParamList.Remove ("@SYS_FncList")
  End If
  
  Dim FncList
  Set FncList = CreateObject("Scripting.Dictionary")
  FncList.CompareMode = 1
  ParamList.Add "@SYS_FncList", FncList

  'Специальные функции только для данного типа отчетов
  ParamList("@SYS_FncList")("Cell") = "Excel_GetCellValue"
  ParamList("@SYS_FncList")("Code128") = "Excel_Code128"
  ParamList("@SYS_FncList")("EAN13") = "Excel_EAN13"
  ParamList("@SYS_FncList")("img") = "Excel_Img"
  
  
  Set WorkBook = CreateObject("Excel.Application").Workbooks.Open(Template)
  WorkBook.SaveAs sOutFile
  
  For Each Sheet In WorkBook.sheets
    Sheet.Activate
    ParamList.Add "@SYS_CurrentSheet", Sheet
    ExcelReportFillSheet Sheet, ExcelReportGetModel(Sheet), ParamList
    'Удалим ссылки которые больше не нужны
    If ParamList.exists("@SYS_PrevRecordset") Then ParamList.Remove "@SYS_PrevRecordset"
    ParamList.Remove ("@SYS_CurrentSheet")
  Next
  
  WorkBook.Application.Visible = True
End Function


Public Function MakeReportXLSX(Template, ParamList, sOutFile, bPrint)
  MakeReportXLSX = MakeReportExcel(Template, ParamList, sOutFile, bPrint)
End Function

Public Function MakeReportXLS(Template, ParamList, sOutFile, bPrint)
  MakeReportXLS = MakeReportExcel(Template, ParamList, sOutFile, bPrint)
End Function

Public Function MakeReportXLSM(Template, ParamList, sOutFile, bPrint)
  MakeReportXLSM = MakeReportExcel(Template, ParamList, sOutFile, bPrint)
End Function

Public Function Excel_GetCellValue(pParamList, aArg As Variant) As String
  If aArg(1) = vbNullString Or aArg(0) = 0 Then
    Excel_GetCellValue = Empty
  Else
    Dim Cell
    Set Cell = pParamList("@SYS_CurrentSheet").Range(aArg(1))
    Excel_GetCellValue = Cell.Value
    If aArg(0) > 1 Then
      Cell.Clear
    End If
    Set Cell = Nothing
  End If
End Function



Public Function Excel_Code128(ByRef pParamList As Object, aArg As Variant) As String
  Excel_Code128 = vbNullString
  On Error GoTo OnError:
  Dim byteStorage() As Byte, BCWidth
  If aArg(0) > 0 And aArg(1) <> vbNullString Then
    byteStorage = StrConv(zebra2wmf(code128_zebra(aArg(1), 3), 2, 40, BCWidth), vbFromUnicode)
    Dim Cell, Image, filename As String
    Set Cell = pParamList("@SYS_CurrentCell")
    filename = "%temp%\picture.emf"
    SaveByteArray byteStorage, filename, True
    Set Image = Cell.Parent.Pictures.Insert(filename)
    Image.Top = Cell.Top
    Image.Left = Cell.Left
    If aArg(0) > 1 Then
      Image.ShapeRange.Width = aArg(2)
    Else
      Image.ShapeRange.Width = Cell.Width
    End If
    If aArg(0) > 2 Then
      Image.ShapeRange.Height = aArg(3)
    Else
      Image.ShapeRange.Height = Cell.Height
    End If
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
End Function


Public Function Excel_EAN13(ByRef pParamList As Object, aArg As Variant) As String
  Excel_EAN13 = vbNullString
  On Error GoTo OnError:
  Dim byteStorage() As Byte, BCWidth
  
  
  If aArg(0) > 0 And aArg(1) <> vbNullString Then
    byteStorage = StrConv(zebra2wmf(EAN13_zebra(aArg(1), False), 2, 40, BCWidth), vbFromUnicode)
    Dim Cell, Image, filename As String
    Set Cell = pParamList("@SYS_CurrentCell")
    filename = "%temp%\picture.emf"
    SaveByteArray byteStorage, filename, True
    Set Image = Cell.Parent.Pictures.Insert(filename)
    Image.Top = Cell.Top
    Image.Left = Cell.Left
    If aArg(0) > 1 Then
      Image.ShapeRange.Width = aArg(2)
    Else
      Image.ShapeRange.Width = Cell.Width
    End If
    If aArg(0) > 2 Then
      Image.ShapeRange.Height = aArg(3)
    Else
      Image.ShapeRange.Height = Cell.Height
    End If
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
End Function

Public Function Excel_Img(ByRef pParamList As Object, aArg As Variant) As String
  Excel_Img = vbNullString
  On Error GoTo OnError:
  Dim byteStorage() As Byte, BCWidth, filename As String
  
  If aArg(0) = 0 Or IsNull(aArg(1)) Or IsEmpty(aArg(1)) Then
    Exit Function
  ElseIf LCase(TypeName(aArg(1))) = "byte()" Then
    filename = "%temp%\picture.emf"
    SaveByteArray aArg(1), filename, True
  Else
    filename = aArg(1)
    If Left(filename, 2) = ".\" Then filename = GetPath(CurrentDb.name) & Mid(filename, 3)
  End If
  
  byteStorage = aArg(1)
  Dim Cell, Image
  Set Cell = pParamList("@SYS_CurrentCell")
  Set Image = Cell.Parent.Pictures.Insert(filename)
  Image.Top = Cell.Top
  Image.Left = Cell.Left
  If aArg(0) > 1 Then
    Image.ShapeRange.Width = aArg(2)
  Else
    Image.ShapeRange.Width = Cell.Width
  End If
  If aArg(0) > 2 Then
    Image.ShapeRange.Height = aArg(3)
  Else
    Image.ShapeRange.Height = Cell.Height
  End If
  
  Exit Function
OnError:
  Dim errNumber, errSource, errDescription: errNumber = Err.Number: errSource = Err.Source: errDescription = Err.Description
  On Error GoTo 0
  Err.Number = errNumber: Err.Source = errSource: Err.Description = errDescription
End Function