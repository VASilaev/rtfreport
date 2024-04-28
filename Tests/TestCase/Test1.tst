TestName: Проверка
TestMethod: ByRunFunction
Function: PrepareRTF 
Params:
 - rawТекст до скана{\field {\*\fldinst {scan("a" for "Клиенты")}{}}}Тело скана{\field {\*\fldinst {endscan()}{}}}Текст после скана
Result: GOTO00000015        PRNT0000000EТекст до сканаOPRS003"a"0009"Клиенты"00000070PRNT0000000AТело сканаNEXT0000004EPRNT00000011Текст после сканаENDT
NextStage:
  Function: MakeReport
  Params:
  - <PrevResult>
  - Nothing
  - Nothing