TestName: ��������
TestMethod: ByRunFunction
Function: PrepareRTF 
Params:
 - raw����� �� �����{\field {\*\fldinst {scan("a" for "�������")}{}}}���� �����{\field {\*\fldinst {endscan()}{}}}����� ����� �����
Result: GOTO00000015        PRNT0000000E����� �� �����OPRS003"a"0009"�������"00000070PRNT0000000A���� �����NEXT0000004EPRNT00000011����� ����� �����ENDT
NextStage:
  Function: MakeReport
  Params:
  - <PrevResult>
  - Nothing
  - Nothing