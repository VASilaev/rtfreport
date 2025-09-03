rem=$cmd$
chcp 1251
set SCRIPT_FILE=%~dpnx0
..\Example\Платежи.accdb /nostartup /x RunFileFromEnv
exit 0
$cmd$

sFile=.\KRNReport.bas
@ExportModule

sFile=.\KRNScripter.bas
@ExportModule

sFile=.\mdQRCodegen.bas
@ExportModule

sFile=.\KRNReportExcel.bas
@ExportModule

sFile=.\PlantFormat.bas
@ExportModule

Save=2
@ApplicationQuit