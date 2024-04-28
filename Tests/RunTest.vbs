msg dumpvar(toKeyValuePairs(ReadFile("d:\Documents\rtfreport\Tests\TestCase\Test1.tst"), nothing))

function dumpVar(byref a)
  on error resume next
  dim i, v

  if isArray(a) then
    if uBound(a) >= lBound(a) then
      dumpVar = dumpVar & "array("
      for i = lbound(a) to ubound(a)
       if i > lBound(a) then
        v = "," & vbCrlf
       else
        v = vbCrlf
       end if
       v = v & ShiftRight(dumpVar(a(i)))
       dumpVar = dumpVar & v
      next
      dumpVar = dumpVar & vbCrlf & ")"
    else
      dumpVar = "array()"
    end if

  elseif TypeName(a) = "Dictionary" then
   dumpVar = "Dic {" & vbCrLf
   for each v in a.keys
    dumpVar = dumpVar & "  " & v & ": " & ShiftRight(dumpVar(a(v))) & vbCrLf
   next
   dumpVar = dumpVar & "}"
  elseif InA(VarType(a), array(2,3,4,5)) then
    dumpVar = a & ""
  elseif TypeName(a) = "JSON" then
    dumpVar = a.dump

  elseif isObject(a)  then
    dumpVar = "object(" & TypeName(a) &")"

  else 'Все остальные как строки
    if a = StartInsert then
     dumpVar = "StartInsert"
    elseif a = EndInsert then
     dumpVar = "EndInsert"
    elseif isNull(a) then
     dumpVar = "[NULL]"     
    else
      dumpVar = """" & replace(a, """", """""") & """"
      if instr(dumpVar,vbCr) > 0 or instr(dumpVar,vbLf) > 0 then
        dumpVar = replace(dumpVar,vbCr,chr(164))
        dumpVar = replace(dumpVar,vblf,chr(182))
        dumpVar = "sf(" & dumpVar & ")"
      end if
    end if

  end if
end function

function ReadFile(fileName)
   dim tf
   Set tf = CreateObject("scripting.FileSystemObject").OpenTextFile(fileName,1,false)
   ReadFile = tf.ReadAll
   tf.Close
   Set tf = Nothing
end function

'Основано на https://github.com/dank8/vba_convert_yaml/tree/master
Function toKeyValuePairs(inputString, dictOfValues)
    'Source: https://stackoverflow.com/questions/38738162/yaml-parser-for-excel-vba
    
    'Replace crlf with cr
    inputString = Replace(inputString, vbCrLf, vbCr)
    'Replace lf with cr
    inputString = Replace(inputString, vbLf, vbCr)
    
    bracesOpen = False
    quoteOpen = False
    keyNameParent = ""
    keyName = ""
    KeyValue = ""
    currentSegment = ""
    Errors = ""

    iLen = Len(inputString)
    if dictOfValues is nothing then  Set dictOfValues = CreateObject("Scripting.Dictionary")
    
    curIndex = 1
    do while curIndex <= iLen
        curCharacter = Mid(inputString, curIndex, 1)
        curIndex = curIndex + 1
    
        Select Case curCharacter
            Case "#"
                commentOpen = True
                do while curIndex <= iLen
                  curCharacter = Mid(inputString, curIndex, 1)
                  curIndex = curIndex + 1
                  if curCharacter = vbCr then exit do
                loop

            Case "{"
                'Ignore braces
                If Not keyName = "" Then
                    keyNameParent = Trim(keyName)
                End If
                keyName = ""
                KeyValue = ""
                currentSegment = ""
            Case "}"
                'Ignore braces
                keyNameParent = ""
            Case "'"
                If quoteOpen Then
                    KeyValue = currentSegment
                    If dictOfValues.exists(keyName) Then
                        Errors = Errors & vbCrLf & "Cannot overwrite existing key."
                    Else
                        If keyNameParent = "" Then
                            dictOfValues.Add keyName, KeyValue
                        Else
                            dictOfValues.Add keyNameParent & "." & keyName, KeyValue
                        End If
                        currentSegment = ""
                        KeyValue = ""
                        keyName = ""
                    End If
                End If
                quoteOpen = Not quoteOpen
            Case ","
                If Not keyName = "" Then
                    KeyValue = Trim(currentSegment)
                    If keyNameParent = "" Then
                        dictOfValues.Add keyName, KeyValue
                    Else
                        dictOfValues.Add keyNameParent & "." & keyName, KeyValue
                    End If
                    keyName = ""
                    KeyValue = ""
                    currentSegment = ""
                End If
                currentSegment = ""
            Case vbCr
                If quoteOpen Then
                    Errors = Errors & vbCrLf & "New line not allowed inside value"
                Else
                    If Not keyName = "" Then
                        KeyValue = Trim(currentSegment)
                        If keyNameParent = "" Then
                            dictOfValues.Add keyName, KeyValue
                        Else
                            dictOfValues.Add keyNameParent & "." & keyName, KeyValue
                        End If
                        keyName = ""
                        KeyValue = ""
                        currentSegment = ""
                    End If
                End If
                currentSegment = ""
            Case vbLf
                'ignore linefeed
            Case ":"
                If quoteOpen Then
                    'Do nothing
                Else
                    keyName = Trim(currentSegment)
                    currentSegment = ""
                End If
            Case Else
                currentSegment = currentSegment & curCharacter
        End Select
    loop
    
    If Not keyName = "" And Not currentSegment = "" Then
        KeyValue = Trim(currentSegment)
        If keyNameParent = "" Then
            dictOfValues.Add keyName, KeyValue
        Else
            dictOfValues.Add keyNameParent & "." & keyName, KeyValue
        End If
        keyName = ""
        KeyValue = ""
        currentSegment = ""
    End If
    
    If Not Errors = "" Then
        dictOfValues.Add "Errors", Errors
    End If
    Set toKeyValuePairs = dictOfValues
End Function


function Msg(sText)
  wscript.echo sText
end function