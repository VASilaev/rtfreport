Текст до скана{\field {\*\fldinst {scan(""a"" for ""Клиенты"")}{}}}Тело скана{\field {\*\fldinst {endscan()}{}}}Текст после скана

GOTO00000015        PRNT0000000EТекст до сканаOPRS003"a"0009"Клиенты"________
GOTO00000015        PRNT0000000EТекст до сканаOPRS003"a"0009"Клиенты"00000070PRNT0000000AТело сканаNEXT0000004EPRNT00000011Текст после
 сканаENDT



?MakeReport(PrepareRTF("rawТекст до скана{\field {\*\fldinst {scan(""a"" for ""Клиенты"")}}{}} {\field {\*\fldinst {f(a.Фамилия)}}{}} {\field {\*\fldinst {endscan()}}{}}Текст после скана"),nothing,nothing)



"rawТекст до скана{\field {\*\fldinst {scan(""a"" for ""Клиенты"")}}{}} {\field {\*\fldinst {f(a.Фамилия)}}{}} {\field {\*\fldinst {endscan()}}{}}Текст после скана"

GOTO00000015        PRNT0000000EТекст до сканаOPRS003"a"0009"Клиенты"00000084PRNT00000001{PRVL009a.ФамилияPRNT00000001}NEXT0000004EPRNT00000011Текст после сканаENDT


"rawТекст до скана{\field {\*\fldinst {scan(""a"" for ""Клиенты"")}}{}}Этот текст будет выведен только на перовй строке{\field {\*\fldinst {scanEntry()}}{}} {\field {\*\fldinst {f(a.Фамилия)}}{}} {\field {\*\fldinst {endscan()}}{}}Текст после скана"

GOTO00000015        PRNT0000000EТекст до сканаOPRS003"a"0009"Клиенты"000000C0PRNT00000030Этот текст будет выведен только на перовй строкеPRNT00000001{PRVL009a.ФамилияPRNT00000001}NEXT0000008APRNT00000011Текст после сканаENDT

"rawТекст до скана{\field {\*\fldinst {scan(""a"" for ""Клиенты"")}}{}}  Entry{\field {\*\fldinst {scanEntry()}}{}} {\field {\*\fldinst {f(a.Фамилия)}}{}} {\field {\*\fldinst {scanfooter()}}{}}  Footer{\field {\*\fldinst {endscan()}}{}}  Текст после скана"

"rawТекст до скана{\field {\*\fldinst {scan(""a"" for ""Клиенты"")}}{}}{\field {\*\fldinst {f(a.Фамилия)}}{}} {\field {\*\fldinst {scanfooter()}}{}}  Footer{\field {\*\fldinst {endscan()}}{}}  Текст после скана"

"rawТекст до скана{\field {\*\fldinst {scan(""a"" for ""select * from Клиенты where 1 = 0"")}}{}}  Entry{\field {\*\fldinst {scanEntry()}}{}} {\field {\*\fldinst {f(a.Фамилия)}}{}} {\field {\*\fldinst {scanfooter()}}{}}  Footer{\field {\*\fldinst {endscan()}}{}}  Текст после скана"


before {\field {\*\fldinst {if(1=0)}}{}}if block{\field {\*\fldinst {endif()}}{}} after

GOTO00000015        PRNT00000007before JMPF0031=00000004EPRNT00000008if blockPRNT00000006 afterENDT

before {\field {\*\fldinst {if(1=0)}}{}}then block{\field {\*\fldinst {else()}}{}}else block{\field {\*\fldinst {endif()}}{}} after

GOTO00000015        PRNT00000007before JMPF0031=00000005CPRNT0000000Athen blockGOTO00000072PRNT0000000Aelse blockPRNT00000006 afterENDT

before {\field {\*\fldinst {if(1=0)}}{}}then block{\field {\*\fldinst {elif(1=0)}}{}}elif1 block{\field {\*\fldinst {else()}}{}}else block{\field {\*\fldinst {endif()}}{}} after

GOTO00000015        PRNT00000007before JMPF0031=00000005CPRNT0000000Athen blockGOTO000000A7JMPF0031=000000091PRNT0000000Belif1 blockGOTO000000A7PRNT0000000Aelse blockPRNT00000006 afterENDT