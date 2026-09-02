Option Compare Database
Option Explicit


' Модуль генерации отчетов в формате Excel из шаблона
' Версия 1.0. 2025 год
' Больше информации на странице https://github.com/VASilaev/rtfreport


Dim CurrentCell, CurrentRow, PrevRecordset, CurrentSheet, CurrentCellFormatterList


Private Const FI_TYPE = 0
Private Const FI_VALUE = 1
Private Const FT_UNKNOWN = -1
Private Const FT_TEXT = 0
Private Const FT_FORMULA = 1


Private Const MD_ROW = 0
Private Const MD_COLUMN = 1
Private Const MD_TYPE = 2


Private Const MD_RECORD_HEIGHT = 3
Private Const MD_RECORD_WIDTH = 4
Private Const MD_RECORD_NAME = 5
Private Const MD_RECORD_SOURCE = 6

Private Const MD_FORMULA_SOURCE = 3

Private Const MDT_RECORD = 0
Private Const MDT_VALUE = 1

Private Const nBlockSize = 1024
 

Public Function ParseTemplateFormula(ByVal sValue)
'Парсим вероятную формулу на элементы
  Dim SlashPos, BreakerPos, Breaker, aFormula, prevItemType, i, UpperBound
  prevItemType = FT_UNKNOWN ' Не известно

    'Парсим выражение, возможно это даже не формула. Формула должна содержать закрытые {}, в качестве символа экранирования используем "\"
  aFormula = Array()
  UpperBound = -1
  Breaker = "{"
  Do While True
    SlashPos = InStr(sValue, "\")
    BreakerPos = InStr(sValue, Breaker)
    If SlashPos = 0 And BreakerPos = 0 Then Exit Do
    If SlashPos = 0 Then SlashPos = Len(sValue) + 1
    If BreakerPos = 0 Then BreakerPos = Len(sValue) + 1
    
    If SlashPos < BreakerPos Then
      If SlashPos > 1 Then
        If prevItemType <> FT_UNKNOWN Then
          aFormula(UpperBound)(FI_VALUE) = aFormula(UpperBound)(FI_VALUE) & Left(sValue, SlashPos - 1)
        Else
          UpperBound = UpperBound + 1
          ReDim Preserve aFormula(UpperBound)
          prevItemType = FT_TEXT
          aFormula(UpperBound) = Array(FT_TEXT, Left(sValue, SlashPos - 1))
        End If
        sValue = Mid(sValue, SlashPos)
      End If
      
      If Left(sValue, 2) = "\\" Or Left(sValue, 2) = "\{" Or Left(sValue, 2) = "\}" Then i = 2 Else i = 1
      If prevItemType <> FT_UNKNOWN Then
        aFormula(UpperBound)(FI_VALUE) = aFormula(UpperBound)(FI_VALUE) & Mid(sValue, i, 1)
      Else
        UpperBound = UpperBound = 1
        ReDim Preserve aFormula(UpperBound)
        prevItemType = FT_TEXT
        aFormula(UpperBound) = Array(FT_TEXT, Mid(sValue, i, 1))
      End If
      sValue = Mid(sValue, i + 1)

    ElseIf Breaker = "{" Then
      If BreakerPos > 1 Then
        If prevItemType <> FT_UNKNOWN Then
          aFormula(UpperBound)(FI_VALUE) = aFormula(UpperBound)(FI_VALUE) & Left(sValue, BreakerPos - 1)
        Else
          UpperBound = UpperBound + 1
          ReDim Preserve aFormula(UpperBound)
          prevItemType = FT_TEXT
          aFormula(UpperBound) = Array(FT_TEXT, Left(sValue, BreakerPos - 1))
        End If
      End If
      sValue = Mid(sValue, BreakerPos + 1)
      
      prevItemType = FT_FORMULA
      UpperBound = UpperBound + 1
      ReDim Preserve aFormula(UpperBound)
      aFormula(UpperBound) = Array(FT_FORMULA, "")
      Breaker = "}"
    Else
      aFormula(UpperBound)(FI_VALUE) = aFormula(UpperBound)(FI_VALUE) & Left(sValue, BreakerPos - 1)
      prevItemType = FT_UNKNOWN
      Breaker = "{"
      sValue = Mid(sValue, BreakerPos + 1)
    End If
  Loop
  
  If Len(sValue) > 0 Then
    If prevItemType = FT_UNKNOWN Then
      UpperBound = UpperBound + 1
      ReDim Preserve aFormula(UpperBound)
      aFormula(UpperBound) = Array(0, sValue)
    ElseIf Breaker = "{" Then
      aFormula(UpperBound)(FI_VALUE) = aFormula(UpperBound)(FI_VALUE) & sValue
    Else 'Формула не закрыта значит это просто текст
      aFormula(UpperBound)(FI_VALUE) = "{" & aFormula(UpperBound)(FI_VALUE) & sValue
      aFormula(UpperBound)(FI_TYPE) = FT_TEXT
      If UpperBound > 0 Then
        If aFormula(UpperBound - 1)(FI_TYPE) = 0 Then
          aFormula(UpperBound - 1)(FI_VALUE) = aFormula(UpperBound - 1)(FI_VALUE) & aFormula(UpperBound)(FI_VALUE)
          UpperBound = UpperBound - 1
          ReDim Preserve aFormula(UpperBound)
        End If
      End If
    End If
  End If

  If UpperBound < 0 Then aFormula = Array(Array(FT_TEXT, ""))

  ParseTemplateFormula = aFormula
End Function

Function ExcelColToLetter(col)
'Преобразует номер колонки в буквенный индекс
'#param col - Номер колонки
  Dim temp, letter
  temp = col
  Do While temp > 0
    letter = Chr(64 + (temp - 1) Mod 26 + 1) & letter
    temp = Int((temp - 1) / 26)
  Loop
  ExcelColToLetter = letter
End Function


Function ExcelReportGetModel(objSheet)
'Извлекаем с листа модель заполнения
  Dim Model, workbookname, rng, UpperBound
  Model = Array()
  UpperBound = -1
  
  workbookname = objSheet.name
  
  
  For Each rng In objSheet.Names
    'Находим именованные диапазоны заданного формата
    If LCase(Right(rng.name, 7)) = ".record" Then
      UpperBound = UpperBound + 1
      ReDim Preserve Model(UpperBound)
      Model(UpperBound) = Array(rng.RefersToRange.Row, rng.RefersToRange.Column, MDT_RECORD, 1, 1, LCase(Mid(rng.name, 1, Len(rng.name) - 7)), rng.Comment)
        
      If IsArray(rng.RefersToRange.Value2) Then
        Model(UpperBound)(MD_RECORD_HEIGHT) = UBound(rng.RefersToRange.Value2, 1)
        Model(UpperBound)(MD_RECORD_WIDTH) = UBound(rng.RefersToRange.Value2, 2)
      End If
    End If
  Next
  
  For Each rng In objSheet.Parent.Names
    'Находим именованные диапазоны заданного формата
    If LCase(Right(rng.name, 7)) = ".record" And Left(rng.RefersTo, Len(workbookname) + 2) = "=" & workbookname & "!" Then
      UpperBound = UpperBound + 1
      ReDim Preserve Model(UpperBound)
      
      Model(UpperBound) = Array(rng.RefersToRange.Row, rng.RefersToRange.Column, MDT_RECORD, 1, 1, LCase(Mid(rng.name, 1, Len(rng.name) - 7)), rng.Comment)
        
      If IsArray(rng.RefersToRange.Value2) Then
        Model(UpperBound)(MD_RECORD_HEIGHT) = UBound(rng.RefersToRange.Value2, 1)
        Model(UpperBound)(MD_RECORD_WIDTH) = UBound(rng.RefersToRange.Value2, 2)
      End If
    End If
  Next
  
  Dim i, j
  
  'Рекрдсеты не должны пересекаться
  For i = LBound(Model) To UpperBound - 1
    For j = i + 1 To UpperBound
      If Model(i)(MD_ROW) < Model(j)(MD_ROW) + Model(j)(MD_RECORD_HEIGHT) And Model(i)(MD_ROW) + Model(i)(MD_RECORD_HEIGHT) > Model(j)(MD_ROW) Then
        Err.Raise 2000, , "Набор данных [" & Model(i)(MD_RECORD_NAME) & "] пересекается с [" & Model(j)(MD_RECORD_NAME) & "]"
      End If
    Next
  Next
  
  'Собираем формулы
  Dim FindedCell, aFormula, StartAddress
  Set FindedCell = objSheet.Cells.Find("*{*}*")
  Do While Not FindedCell Is Nothing
    If IsEmpty(StartAddress) Then StartAddress = FindedCell.Address Else If StartAddress = FindedCell.Address Then Exit Do
    
    aFormula = ParseTemplateFormula(FindedCell.FormulaR1C1)
      
    'Если не формула - игнорируем
    If Not (UBound(aFormula) = 0 And aFormula(0)(FI_TYPE) = FT_TEXT) Then
      UpperBound = UpperBound + 1
      ReDim Preserve Model(UpperBound)
      Model(UpperBound) = Array(FindedCell.Row, FindedCell.Column, MDT_VALUE, aFormula)
    End If
    Set FindedCell = objSheet.Cells.FindNext(FindedCell)
  Loop
  
  'Сортируем элементы
  Dim swap
  For i = LBound(Model) To UpperBound - 1
    For j = i + 1 To UpperBound
      swap = False
            
      If Model(i)(MD_TYPE) = MDT_RECORD And Model(j)(MD_TYPE) = MDT_VALUE Then
        If Model(j)(MD_ROW) >= Model(i)(MD_ROW) And Model(j)(MD_ROW) < Model(i)(MD_ROW) + Model(i)(MD_RECORD_HEIGHT) And (Model(j)(MD_COLUMN) < Model(i)(MD_COLUMN) Or Model(j)(MD_COLUMN) >= Model(i)(MD_COLUMN) + Model(i)(MD_RECORD_WIDTH)) Then
          'Специальный случай, формулы которые попали в строки рекордсета, но находятся вне его диапазона по столбцам должны обработаться до самого рекордсета
          swap = True
        End If
      End If
      
      If Not swap And (Model(j)(MD_ROW) < Model(i)(MD_ROW) Or _
             (Model(j)(MD_ROW) = Model(i)(MD_ROW) And Model(j)(MD_COLUMN) < Model(i)(MD_COLUMN)) Or _
             (Model(j)(MD_ROW) = Model(i)(MD_ROW) And Model(j)(MD_COLUMN) = Model(i)(MD_COLUMN) And Model(j)(MD_TYPE) < Model(i)(MD_TYPE))) Then
        'Заполняемые ячейки заполняются сверху вниз, слева направо. В первую очередь обрабатывается рекордсет.
        swap = True
      End If
      
      If swap Then
        swap = Model(i)
        Model(i) = Model(j)
        Model(j) = swap
      End If
    Next
  Next
  
  ExcelReportGetModel = Model
End Function

Public Sub ExcelAddCellFormatter(ByRef ParamList, sProcFormatter, pUserData)

'Регистрирует форматирование ячейки
'#param ParamList - Текущий контекст
'#Param sProcFormatter - имя функции форматирования имеет следующие параметры:
' pRange - диапазон текущей редактируемой ячейки
' ParamList - текущий контекст
' pUserData - копия данных переданных в ExcelAddCellFormatter
'#param pUserData - пользовательские данные будут переданы при вызове форматтера

KRNReport.addInArray CurrentCellFormatterList, Array(sProcFormatter, pUserData)

End Sub

Private Function ExcelReportCalcCellValue(aFormula, ParamList)
  Dim FormulaValue, FormulaElement

  'Заполнение формул вида {формула}
  FormulaValue = Empty
  
  CurrentCellFormatterList = Array()
  
  For Each FormulaElement In aFormula
    If FormulaElement(FI_TYPE) = FT_TEXT Then
      FormulaValue = FormulaValue & FormulaElement(FI_VALUE)
    ElseIf FormulaElement(FI_TYPE) = FT_FORMULA Then
      If IsEmpty(FormulaValue) Then
        FormulaValue = GetExpression(FormulaElement(FI_VALUE), ParamList)
      Else
        FormulaValue = FormulaValue & GetExpression(FormulaElement(FI_VALUE), ParamList)
      End If
    Else
      Err.Raise 2002, , "Что то пошло не так модель сломалась"
    End If
  Next

  ExcelReportCalcCellValue = Array(FormulaValue, CurrentCellFormatterList)
  CurrentCellFormatterList = Array()
End Function

Private Sub ExcelReportFormatCell(pCell, pFormatterList, ParamList)
  Dim FncFormatter, sErrorMsg
  Set CurrentCell = pCell
  
  If ParamList.Exists("@SYS_CurrentCell") Then ParamList.Remove ("@SYS_CurrentCell")
  ParamList.Add "@SYS_CurrentCell", pCell
  On Error Resume Next
  For Each FncFormatter In pFormatterList
    Application.Run FncFormatter(0), CurrentCell, ParamList, FncFormatter(1)
    
    If Err.Number <> 0 Then
      sErrorMsg = "Ошибка в формуле {" & CurrentCell.Value & "} ячейки (" & CurrentCell.Row & "," & CurrentCell.Column & ") при обработке форматтером " & FncFormatter(0) & "." & vbCrLf & "[" & Err.Number & "]" & Err.Description
      Err.Clear
      On Error GoTo 0
      Err.Raise 1000, , sErrorMsg
    End If
  Next
  On Error GoTo 0
  Set CurrentCell = Nothing
  ParamList.Remove "@SYS_CurrentCell"
End Sub

Public Function ExcelReportFillSheet(objSheet, aModel, ParamList, Optional ByVal nRowStart = -1, Optional ByVal nRowEnd = -1)
  Dim i, j, aRecordSet, sErrorMsg, CellRange, row1, row2, col1, col2, RecordResult, RecordValues, RecordValuesEtalon, RecordFormats, nBlockSizeLocal
  Dim CurrentOffsetRow, ElementsInDataset, bShift, RecordHeight
  
  If nRowStart = -1 Then nRowStart = LBound(aModel)
  If nRowEnd = -1 Then nRowEnd = UBound(aModel)
  
  For i = nRowStart To nRowEnd
    Select Case aModel(i)(MD_TYPE)
      Case MDT_VALUE
        RecordResult = ExcelReportCalcCellValue(aModel(i)(MD_FORMULA_SOURCE), ParamList)
           
        Set CurrentCell = objSheet.Cells(aModel(i)(MD_ROW), aModel(i)(MD_COLUMN))
        CurrentCell.FormulaR1C1 = RecordResult(0)
        
        ExcelReportFormatCell CurrentCell, RecordResult(1), ParamList
        
      Case MDT_RECORD
        'Обработка наборов данных, вложенные отсекаются на уровне модели
        
        CurrentOffsetRow = 0
        ElementsInDataset = 0
        RecordHeight = aModel(i)(MD_RECORD_HEIGHT)
        
        'Простая таблица без вложенных подтаблиц, ее можно заполнить одним массивом.
        Dim bIsSimpleTable
        bIsSimpleTable = True
        
        For j = i + 1 To UBound(aModel)
          If aModel(j)(MD_ROW) >= aModel(i)(MD_ROW) + aModel(i)(MD_RECORD_HEIGHT) Then Exit For
          If aModel(j)(MD_TYPE) <> MDT_VALUE Then bIsSimpleTable = False
          ElementsInDataset = ElementsInDataset + 1
        Next
        
        'OpenRecordsetForReport возвращает EOF = True если нет строк с данными
        If Not OpenRecordsetForReport(aModel(i)(MD_RECORD_NAME), aModel(i)(MD_RECORD_SOURCE), ParamList, aRecordSet) Then
        
          row1 = aModel(i)(MD_ROW)
          row2 = row1 + RecordHeight - 1
          col1 = aModel(i)(MD_COLUMN)
          col2 = col1 + aModel(i)(MD_RECORD_WIDTH) - 1
          
          If bIsSimpleTable Then
            ' =========================================================================
            ' БЛОЧНЫЙ АЛГОРИТМ ЗАПОЛНЕНИЯ (ПРОСТАЯ ТАБЛИЦА)
            ' =========================================================================
            Dim nBlockRows, nAllocatedRecords, nCurrentRecordInBlock, nTotalProcessedRecords
            Dim rBlock1, rBlock2, bInitialized, singleEtalon
            Dim CellFormulaIdx, CellRelRow, CellRelCol
            Dim flushRng, tailRowsToDelete, unscaledData
            
            
            nBlockSizeLocal = KRNReport.RecordCountRecordsetForReport(ParamList, aRecordSet)
            If IsEmpty(nBlockSizeLocal) Then
              nBlockSizeLocal = nBlockSize
            Else
              nBlockSizeLocal = nBlockSizeLocal \ 4
              If nBlockSizeLocal = 0 Then
                nBlockSizeLocal = 1
              ElseIf nBlockSizeLocal > nBlockSize Then
                nBlockSizeLocal = nBlockSize
              End If
            End If
            
            nBlockRows = nBlockSizeLocal * RecordHeight
            nAllocatedRecords = 0
            nCurrentRecordInBlock = 0
            nTotalProcessedRecords = 0
            bInitialized = False
            RecordFormats = Array()
            
            ' Начало текущего блока выгрузки
            rBlock1 = row1
            
            
            
            Do While FetchRow(ParamList)
              ' --- 1. Первичная инициализация при первой строке набора данных ---
              If Not bInitialized Then
                ' Снимаем эталон одной записи (1..RecordHeight, 1..RecordWidth)
                Set CurrentRow = objSheet.Range(ExcelColToLetter(col1) & row1 & ":" & ExcelColToLetter(col2) & row2)
                If IsArray(CurrentRow.FormulaR1C1) Then
                  singleEtalon = CurrentRow.FormulaR1C1
                Else
                  If IsEmpty(singleEtalon) Then singleEtalon = Array()
                  ReDim singleEtalon(1 To 1, 1 To 1)
                  singleEtalon(1, 1) = CurrentRow.FormulaR1C1
                End If
                Set CurrentRow = Nothing
                
                ' Формируем эталонный массив на весь блок nBlockSizeLocal записей
                ReDim RecordValuesEtalon(1 To nBlockRows, 1 To aModel(i)(MD_RECORD_WIDTH))
                Dim b_k, b_r, b_c
                For b_k = 0 To nBlockSizeLocal - 1
                  For b_r = 1 To RecordHeight
                    For b_c = 1 To aModel(i)(MD_RECORD_WIDTH)
                      RecordValuesEtalon(b_k * RecordHeight + b_r, b_c) = singleEtalon(b_r, b_c)
                    Next
                  Next
                Next
                
                ' Раздвигаем Excel сразу до nBlockSize * 2 логических записей (с учетом исходного шаблона)
                ' Вставка происходит ПЕРЕД текущей позицией row1, чтобы формулы снизу расширялись
                Dim nTargetRecords, nCurrentCopies, nToCopy, nInsertedHeight
                
                nTargetRecords = nBlockSizeLocal * 2
                nCurrentCopies = 1
                
                Do While nCurrentCopies < nTargetRecords
                  nToCopy = nCurrentCopies
                  If nCurrentCopies + nToCopy > nTargetRecords Then
                    nToCopy = nTargetRecords - nCurrentCopies
                  End If
                  
                  ' Копируем уже накопленный блок строк
                  Set CellRange = objSheet.Rows(row1 & ":" & (row1 + nToCopy * RecordHeight - 1))
                  CellRange.Copy
                  
                  ' Вставляем ПЕРЕД текущей строкой row1, сдвигая всё (включая формулы) вниз
                  objSheet.Rows(row1).Insert (-4121) ' xlDown
                  objSheet.Application.CutCopyMode = False
                  Set CellRange = Nothing
                  
                  ' Смещаем указатель начала блока вверх на высоту вставки
                  'row1 = row1 + nToCopy * RecordHeight
                  nCurrentCopies = nCurrentCopies + nToCopy
                Loop
                
                nAllocatedRecords = nBlockSizeLocal * 2
                RecordValues = RecordValuesEtalon
                bInitialized = True
              End If
              
              ' --- 2. Расчет значений ячеек для текущей записи ---
              For CellFormulaIdx = i + 1 To i + ElementsInDataset
                CellRelRow = aModel(CellFormulaIdx)(MD_ROW) - aModel(i)(MD_ROW)
                CellRelCol = aModel(CellFormulaIdx)(MD_COLUMN) - aModel(i)(MD_COLUMN)
                
                RecordResult = ExcelReportCalcCellValue(aModel(CellFormulaIdx)(MD_FORMULA_SOURCE), ParamList)
                
                ' Запись в текущую позицию внутри блока
                RecordValues(nCurrentRecordInBlock * RecordHeight + CellRelRow + 1, CellRelCol + 1) = RecordResult(0)
                
                ' Сохраняем постобработчик с абсолютными координатами листа
                If UBound(RecordResult(1)) >= 0 Then
                  ReDim Preserve RecordFormats(UBound(RecordFormats) + 1)
                  RecordFormats(UBound(RecordFormats)) = Array(rBlock1 + nCurrentRecordInBlock * RecordHeight + CellRelRow, col1 + CellRelCol, RecordResult(1))
                End If
              Next
              
              nCurrentRecordInBlock = nCurrentRecordInBlock + 1
              nTotalProcessedRecords = nTotalProcessedRecords + 1
              
              ' Сброс полного блока при достижении nBlockSizeLocal
              If nCurrentRecordInBlock = nBlockSizeLocal Then
                rBlock2 = rBlock1 + nBlockRows - 1
                Set flushRng = objSheet.Range(ExcelColToLetter(col1) & rBlock1 & ":" & ExcelColToLetter(col2) & rBlock2)
                flushRng.FormulaR1C1 = RecordValues
                Set flushRng = Nothing
                
                ' Выполнение постобработчиков накопленного блока
                For Each RecordResult In RecordFormats
                  ExcelReportFormatCell objSheet.Cells(RecordResult(0), RecordResult(1)), RecordResult(2), ParamList
                Next
                RecordFormats = Array()
                
                ' Смещение указателя начала следующего блока вверх
                rBlock1 = rBlock1 + nBlockRows
                rBlock2 = rBlock1 + nBlockRows - 1
                
                ' Пополнение буфера: вставляем новый блок ПЕРЕД текущим rBlock1
                Set CellRange = objSheet.Rows(rBlock1 & ":" & rBlock2)
                CellRange.Copy
                objSheet.Rows(rBlock1).Insert (-4121) ' xlDown
                objSheet.Application.CutCopyMode = False
                Set CellRange = Nothing
                
                nAllocatedRecords = nAllocatedRecords + nBlockSizeLocal
                
                nCurrentRecordInBlock = 0
                RecordValues = RecordValuesEtalon
              End If
            Loop
            
            
            ' --- 4. Финализация: сброс остатка и очистка хвоста ---
            If bInitialized Then
              ' Сброс неполного остатка (если есть)
              If nCurrentRecordInBlock > 0 Then
                Dim nTailRows
                nTailRows = nCurrentRecordInBlock * RecordHeight
                rBlock2 = rBlock1 + nTailRows - 1
                
                ReDim unscaledData(1 To nTailRows, 1 To aModel(i)(MD_RECORD_WIDTH))
                For b_r = 1 To nTailRows
                  For b_c = 1 To aModel(i)(MD_RECORD_WIDTH)
                    unscaledData(b_r, b_c) = RecordValues(b_r, b_c)
                  Next
                Next
                
                Set flushRng = objSheet.Range(ExcelColToLetter(col1) & rBlock1 & ":" & ExcelColToLetter(col2) & rBlock2)
                flushRng.FormulaR1C1 = unscaledData
                Set flushRng = Nothing
                
                For Each RecordResult In RecordFormats
                  ExcelReportFormatCell objSheet.Cells(RecordResult(0), RecordResult(1)), RecordResult(2), ParamList
                Next
                RecordFormats = Array()
                
                ' rBlock1 теперь указывает на начало последнего блока, а rBlock2 - на его конец
                ' Следующий блок (резерв) начинается с rBlock2 + 1
              End If
              
              ' Удаление лишнего зарезервированного хвоста строк (после фактических данных)
              tailRowsToDelete = (nAllocatedRecords - nTotalProcessedRecords) * RecordHeight
              If tailRowsToDelete > 0 Then
                ' Фактический конец данных: rBlock2 (если был остаток) или rBlock1 - RecordHeight (если остатка не было)
                Dim actualEndRow As Long
                If nCurrentRecordInBlock > 0 Then
                  actualEndRow = rBlock1 + (nCurrentRecordInBlock * RecordHeight) - 1
                Else
                  ' Если остатка не было, последний блок был полным и его конец = rBlock1 - RecordHeight
                  actualEndRow = rBlock1 - RecordHeight
                End If
                
                ' Удаляем строки начиная с actualEndRow + 1
                objSheet.Rows(actualEndRow + 1 & ":" & (actualEndRow + tailRowsToDelete)).Delete (-4162) ' xlUp
              End If
              
              ' Сброс UsedRange
              Dim ur
              Set ur = objSheet.UsedRange
              Set ur = Nothing
              
              CurrentOffsetRow = nTotalProcessedRecords * RecordHeight
              
              ' Сдвиг координат последующих элементов модели на фактически добавленную высоту
              Dim nAddedHeight
              nAddedHeight = (nTotalProcessedRecords - 1) * RecordHeight
              If nAddedHeight <> 0 Then
                For j = i + 1 To UBound(aModel)
                  aModel(j)(MD_ROW) = aModel(j)(MD_ROW) + nAddedHeight
                Next
              End If
            End If
            
          Else
            ' =========================================================================
            ' ИСХОДНЫЙ ПОСТРОЧНЫЙ АЛГОРИТМ (СЛОЖНАЯ ТАБЛИЦА С ВЛОЖЕНИЯМИ)
            ' =========================================================================
            RecordValuesEtalon = Empty
            
            Do While FetchRow(ParamList)
              If Not EOFRecordsetForReport(ParamList, aRecordSet) Then
                Set CellRange = objSheet.Rows(row1 & ":" & row2)
                CellRange.Copy
                CellRange.Insert (-4121) ' xlDown
                Set CellRange = Nothing
                bShift = True
              Else
                bShift = False
              End If
              
              Set CurrentRow = objSheet.Range(ExcelColToLetter(col1) & row1 & ":" & ExcelColToLetter(col2) & row2)
              
              ' Рекурсивное заполнение дочерних таблиц/данных
              If ElementsInDataset > 0 Then ExcelReportFillSheet objSheet, aModel, ParamList, i + 1, i + ElementsInDataset
              
              If bShift Then
                For j = i + 1 To UBound(aModel)
                  aModel(j)(MD_ROW) = aModel(j)(MD_ROW) + RecordHeight
                Next
              End If
              
              CurrentOffsetRow = CurrentOffsetRow + RecordHeight
              row1 = row1 + RecordHeight
              row2 = row2 + RecordHeight
            Loop
          End If
          
        Else
          ' Удаляем шаблон (если набор данных пуст)
          row1 = aModel(i)(MD_ROW)
          row2 = row1 + RecordHeight - 1
          objSheet.Rows(row1 & ":" & row2).Delete (-4162) ' xlUp
          ' Сдвигаем в обратном направлении
          For j = i + 1 To UBound(aModel)
            aModel(j)(MD_ROW) = aModel(j)(MD_ROW) - RecordHeight
          Next
        End If
        
        Set CurrentRow = Nothing
       
        ' Поместим в контекст обработанный диапазон
        If CurrentOffsetRow = 0 Then
          Set PrevRecordset = Nothing
        Else
          row1 = aModel(i)(MD_ROW)
          row2 = row1 + CurrentOffsetRow - 1
          col1 = aModel(i)(MD_COLUMN)
          col2 = col1 + aModel(i)(MD_RECORD_WIDTH) - 1
          Set PrevRecordset = objSheet.Range(ExcelColToLetter(col1) & row1 & ":" & ExcelColToLetter(col2) & row2)
        End If
        
        CloseRecordsetForReport ParamList, aRecordSet
        
        i = i + ElementsInDataset
      Case Else
        Err.Raise 2001, , "Что то пошло не так модель сломалась"
    End Select
  Next
  
  Exit Function

OnError:
If Err.Number = 2001 Then
  sErrorMsg = Err.Description
Else
  sErrorMsg = "Ошибка при заполнении листа Excel" & vbCrLf & Err.Description & vbCrLf & vbCrLf & DumpContext(ParamList)
End If
On Error GoTo 0
Err.Clear
Resume ResumeOnError
ResumeOnError:
Err.Raise 2001, , sErrorMsg
End Function


Private Function MakeReportExcel(Template, ParamList, sOutFile, bPrint)
  Dim WorkBook, Sheet, Excel, sError
  
  AddSpecialFunction ParamList, "Excel_GetCellValue", "Cell"
  AddSpecialFunction ParamList, "Excel_Code128", "Code128"
  AddSpecialFunction ParamList, "Excel_EAN13", "EAN13"
  AddSpecialFunction ParamList, "Excel_Img", "img"
  Set Excel = CreateObject("Excel.Application")
  On Error GoTo CloseExcel
  Set WorkBook = Nothing
  Set WorkBook = Excel.Workbooks.Open(Template)
  WorkBook.SaveAs sOutFile
  
  Excel.ScreenUpdating = False
  Excel.EnableEvents = False
  Excel.DisplayAlerts = False
  Excel.Visible = True
  
  For Each CurrentSheet In WorkBook.sheets
    CurrentSheet.Activate
    ExcelReportFillSheet CurrentSheet, ExcelReportGetModel(CurrentSheet), ParamList
    'Удалим ссылки которые больше не нужны
    Set PrevRecordset = Nothing
  Next
  
  Set CurrentSheet = Nothing
  
  Excel.DisplayAlerts = True
  Excel.EnableEvents = True
  Excel.ScreenUpdating = True
  Excel.Visible = True
  Exit Function

ExitFunction:
  Exit Function

CloseExcel:
  sError = Err.Description
  Err.Clear
  Excel.DisplayAlerts = True
  Excel.EnableEvents = True
  Excel.ScreenUpdating = True
  Excel.Visible = True
  WorkBook.Close False
  Excel.Quit
  MsgBox "При заполнении шаблона произошла ошибка" & vbCrLf & sError, vbCritical + vbOKOnly, "Заполнение шаблона Excel"
  
  Resume ExitFunction
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
    Set Cell = CurrentSheet.Range(aArg(1))
    Excel_GetCellValue = Cell.Value
    If aArg(0) > 1 Then
      Cell.Clear
    End If
    Set Cell = Nothing
  End If
End Function

Private Function InsertImgIntoCell(pCell, pFilename, pWidth, pHeight, pAddParam)
  Dim Image, vWidth
  Set Image = pCell.Parent.Pictures.Insert(pFilename)
  Image.Top = pCell.Top
  Image.Left = pCell.Left
  If IsEmpty(pWidth) Then vWidth = pCell.width Else vWidth = pWidth
    
  Image.ShapeRange.width = vWidth
  
  If IsEmpty(pHeight) Then
    Image.ShapeRange.Height = pCell.Height
  Else
    Image.ShapeRange.Height = pHeight
  End If

  If Image.ShapeRange.width > vWidth Then Image.ShapeRange.width = vWidth
End Function

Public Function Excel_Code128(ByRef pParamList As Object, aArg As Variant) As String
  Excel_Code128 = vbNullString
  On Error GoTo OnError:
  Dim byteStorage() As Byte, BCWidth
  If aArg(0) > 0 And aArg(1) <> vbNullString Then
    byteStorage() = StrConv(zebra2wmf(code128_zebra(aArg(1), 3), 2, 40, BCWidth), vbFromUnicode)
    aArg(1) = byteStorage
    Excel_Code128 = Excel_Img(pParamList, aArg)
  End If
  
  Exit Function
OnError:
  Dim ErrNumber, ErrSource, ErrDescription: ErrNumber = Err.Number: ErrSource = Err.Source: ErrDescription = Err.Description
  On Error GoTo 0
  Err.Number = ErrNumber: Err.Source = ErrSource: Err.Description = ErrDescription
End Function


Public Function Excel_EAN13(ByRef pParamList As Object, aArg As Variant) As String
  Excel_EAN13 = vbNullString
  On Error GoTo OnError:
  Dim byteStorage() As Byte, BCWidth, vWidth, vHeight
  
  
  If aArg(0) > 0 And aArg(1) <> vbNullString Then
    byteStorage() = StrConv(zebra2wmf(EAN13_zebra(aArg(1), False), 2, 40, BCWidth), vbFromUnicode)
    aArg(1) = byteStorage
    Excel_EAN13 = Excel_Img(pParamList, aArg)
  End If
  
  Exit Function
OnError:
  Dim ErrNumber, ErrSource, ErrDescription: ErrNumber = Err.Number: ErrSource = Err.Source: ErrDescription = Err.Description
  On Error GoTo 0
  Err.Number = ErrNumber: Err.Source = ErrSource: Err.Description = ErrDescription
End Function

Public Function Excel_Img_PostProcess(ByRef pRange, ByRef pParamList, aArg As Variant)
  Excel_Img_PostProcess = vbNullString
  On Error GoTo OnError:
  Dim BCWidth, filename As String, vWidth, vHeight
  
  If aArg(0) = 0 Then
    Exit Function
  ElseIf IsNull(aArg(1)) Or IsEmpty(aArg(1)) Then
    Exit Function
  ElseIf LCase(TypeName(aArg(1))) = "byte()" Then
    filename = "%temp%\picture."
    Dim vExt
    vExt = LCase(GetTypeContent(aArg(1)))
    Select Case vExt
      Case "jpg", "png", "emf", "wmf":
        filename = filename & vExt
      Case Else
        filename = filename & "unk"
    End Select
    SaveByteArray aArg(1), filename, True
  Else
    filename = aArg(1)
    If Left(filename, 2) = ".\" Then filename = GetPath(CurrentDb.name) & Mid(filename, 3)
  End If
  
  If aArg(0) > 1 Then vWidth = aArg(2) Else vWidth = Empty
  If aArg(0) > 2 Then vHeight = aArg(3) Else vHeight = Empty
  InsertImgIntoCell pRange, filename, vWidth, vHeight, Empty

  Exit Function
OnError:
  Dim ErrNumber, ErrSource, ErrDescription: ErrNumber = Err.Number: ErrSource = Err.Source: ErrDescription = Err.Description
  On Error GoTo 0
  Err.Number = ErrNumber: Err.Source = ErrSource: Err.Description = ErrDescription
End Function

Public Function Excel_Img(ByRef pParamList As Object, aArg As Variant) As String
  ExcelAddCellFormatter pParamList, "Excel_Img_PostProcess", aArg
End Function