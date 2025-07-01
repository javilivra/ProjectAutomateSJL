' Módulo: mod_InstrumentacionAutomatica
Option Explicit

Sub ExtraerAtributosBloqueInstrumentos()
    ' === VARIABLES PARA AUTOCAD ===
    Dim acadApp As Object
    Dim acadDoc As Object
    Dim modelSpace As Object
    Dim entidad As Object
    Dim bloqueRef As Object
    Dim arrAttribs As Variant
    Dim atributo As Object

    ' === VARIABLES PARA EXCEL ===
    Dim wb As Workbook
    Dim hojaLI As Worksheet
    Dim hojaCar As Worksheet
    Dim hojaNotasRef As Worksheet

    ' === VARIABLES DE CONTROL Y PROCESO ===
    Dim valorFunction As String
    Dim valorTag As String
    Dim filaExcel As Long
    Dim tagTracking As Object
    Dim cantidadDuplicados As Long
    Dim listaDuplicados As String

    ' Declaración y asignación separadas para mayor claridad
    Set tagTracking = CreateObject("Scripting.Dictionary")
    cantidadDuplicados = 0
    listaDuplicados = ""

    ' === VARIABLES PARA "LISTA DE DOCUMENTOS" ===
    Dim rutaDocs As Variant
    Dim wbDocs As Workbook
    Dim hojaDocs As Worksheet
    Dim lastRowDocs As Long
    Dim iRowDocs As Long
    Dim matchRow As Long
    Dim codeAES As Variant
    Dim codeYPFProj As Variant
    Dim descDoc As Variant
    Dim codPID As Variant
    Dim posVCD As Long
    Dim vcdCode As String

Set wb = ActiveWorkbook
Set hojaLI = wb.Worksheets("LI")
Set hojaCar = wb.Worksheets("Carátula")
Set hojaNotasRef = wb.Worksheets("Notas - Referencias")

rutaDocs = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Selecciona LISTA DE DOCUMENTOS")
If rutaDocs = False Then Exit Sub
Set wbDocs = Workbooks.Open(rutaDocs, ReadOnly:=True)
Set hojaDocs = wbDocs.Sheets(1)
lastRowDocs = hojaDocs.Cells(hojaDocs.Rows.Count, "C").End(xlUp).Row

' Cargar todos los datos relevantes en un array y cerrar el archivo externo
Dim datosDocs As Variant
datosDocs = hojaDocs.Range("A1:C" & lastRowDocs).Value

wbDocs.Close SaveChanges:=False
Set wbDocs = Nothing
Set hojaDocs = Nothing

MsgBox "Por favor, espera mientras se extraen los datos desde AutoCAD. Este proceso puede tardar varios segundos seg�n la cantidad de bloques."

matchRow = 0
For iRowDocs = 2 To lastRowDocs
    If InStr(1, UCase(datosDocs(iRowDocs, 3)), "LISTA DE INSTRUMENTOS") > 0 Then
        matchRow = iRowDocs
        Exit For
    End If
Next iRowDocs
If matchRow = 0 Then
    MsgBox "No se encontró 'LISTA DE INSTRUMENTOS' en la columna C.", vbExclamation
    Exit Sub
End If

codeAES = datosDocs(matchRow, 1)
codeYPFProj = datosDocs(matchRow, 2)
descDoc = datosDocs(matchRow, 3)
codPID = datosDocs(2, 2)

posVCD = InStr(1, codeYPFProj, "VCD", vbTextCompare)
If posVCD > 0 Then
    vcdCode = Mid(codeYPFProj, posVCD, 8)
Else
    vcdCode = ""
End If

With hojaCar
    .Range("M4").Value = codeAES
    .Range("M2").Value = codeYPFProj
    .Range("B11").Value = descDoc
    .Range("B8").Value = vcdCode
End With

With hojaLI
    .Range("F7").ClearContents
    .Range("F7").Value = codPID
End With

With hojaNotasRef
    .Range("B36").Value = codeYPFProj
    .Range("J36").Value = "P&ID - (Completar con nombre del documento)"
End With

Dim typeMap As Object
Set typeMap = CreateObject("Scripting.Dictionary")
With typeMap
        .Add "AE", "Sonda de analizador":   .Add "AIT", "Analizador Indicador Transmisor"
        .Add "AI", "Indicaci�n de lectura anal�tica": .Add "AS", "Contacto o se�al discreta"
        .Add "AX", "Contacto o se�al discreta": .Add "AL", "Contacto o se�al discreta"
        .Add "FE", "Elemento primario de Caudal": .Add "FG", "Visor en l�nea"
        .Add "FIT", "Caudal�metro": .Add "FI", "Indicaci�n de Caudal"
        .Add "FS", "Contacto o se�al discreta": .Add "FX", "Funci�n/C�lculo"
        .Add "FF", "Relaci�n de caudales": .Add "RO", "Orificio de restricci�n"
        .Add "FIC", "Indicador Controlador de Caudal": .Add "LE", "Elemento primario de medida de nivel"
        .Add "LG", "Nivel visual de vidrio": .Add "LI", "Indicaci�n de Nivel"
        .Add "LIT", "Transmisor de Nivel": .Add "LS", "Contacto o se�al discreta"
        .Add "LX", "Funci�n/C�lculo": .Add "LIC", "Indicador Controlador de Nivel"
        .Add "PI", "Man�metro": .Add "PDI", "Man�metro Diferencial"
        .Add "PIT", "Transmisor de presi�n": .Add "PDIT", "Transmisor de presi�n Diferencial"
        .Add "PS", "Switch de Presi�n": .Add "PDS", "Switch de Presi�n Diferencial"
        .Add "PIC", "Indicador Controlador de Presi�n": .Add "TP", "Prueba de temperatura"
        .Add "TW", "Termovaina": .Add "TE", "Sensor de temperatura"
        .Add "TI", "Term�metro": .Add "TDI", "Term�metro diferencial"
        .Add "TIT", "Transmisor de temperatura": .Add "TS", "Contacto o se�al discreta"
        .Add "TX", "Contacto o se�al discreta": .Add "TL", "Contacto o se�al discreta"
        .Add "TIC", "Indicador Controlador de Temperatura": .Add "AV", "V�lvula de control"
        .Add "FV", "V�lvula de control": .Add "HV", "V�lvula de control"
        .Add "LV", "V�lvula de control": .Add "PV", "V�lvula de control"
        .Add "PDV", "V�lvula de control": .Add "TV", "V�lvula de control"
        .Add "XV", "V�lvula ON-OFF": .Add "SDV", "V�lvula Shutdown"
        .Add "BDV", "V�lvula Blowdown": .Add "MOV", "V�lvula Motorizada"
        .Add "LBV", "V�lvula de corte de ductos": .Add "LCV", "V�lvula autorreguladora por nivel"
        .Add "PCV", "V�lvula autorreguladora por presi�n": .Add "PDCV", "V�lvula autorreguladora por presi�n diferencial"
        .Add "TCV", "V�lvula autorreguladora por temperatura": .Add "SV", "V�lvula solenoide"
        .Add "ZS", "Switch de posici�n": .Add "ZT", "Transmisor de posici�n"
        .Add "ZSO", "Interruptor de posici�n V�lvula abierta": .Add "ZSC", "Interruptor de posici�n V�lvula cerrada"
        .Add "ZLO", "Indicaci�n en pantalla de v�lvula abierta": .Add "ZLC", "Indicaci�n en pantalla de v�lvula cerrada"
        .Add "PSE", "Disco de ruptura": .Add "PSV", "V�lvula de seguridad/alivio"
        .Add "PVSV", "V�lvula de presi�n/vac�o": .Add "HS", "Pulsador"
        .Add "XL", "L�mpara": .Add "YL", "L�mpara"
        .Add "XA", "Alarma": .Add "XSMP", "Orden de marcha/paro"
        .Add "XSM", "Orden de marcha": .Add "XSP", "Orden de paro"
        .Add "XSE", "Permisivo de arranque": .Add "XSB", "Orden de disparo/se�al de bloqueo"
        .Add "XSA", "Orden de abrir V�lvula motorizada": .Add "XSC", "Orden de cerrar V�lvula motorizada"
        .Add "XSD", "Orden de detener V�lvula motorizada": .Add "XY", "-"
        .Add "YM", "Confirmaci�n de marcha": .Add "YR", "Mando en remoto"
        .Add "YD", "Confirmaci�n de equipo disponible": .Add "YS", "-"
        .Add "YA", "Estado de Falla": .Add "XST", "Consigna velocidad o frecuencia"
        .Add "XZT", "Consigna de posici�n": .Add "XET", "Consigna de tensi�n"
        .Add "XIT", "Consigna de intensidad": .Add "XJT", "Consigna de potencia"
        .Add "XGT", "Consigna de cos": .Add "XYT", "Otra variable a especificar"
        .Add "ST", "Velocidad o frecuencia": .Add "ET", "Tensi�n"
        .Add "IT", "Intensidad": .Add "JT", "Potencia"
        .Add "GT", "Cos": .Add "YT", "Otra variable a especificar"
        .Add "BE", "Detector de llama": .Add "BT", "Detector de llama"
        .Add "BI", "Indicaci�n de llama": .Add "BS", "Contacto o se�al discreta"
        .Add "BL", "Estado detector": .Add "SE", "Sonda de medida de velocidad"
        .Add "SS", "Contacto o se�al discreta": .Add "VE", "Sonda de vibraci�n"
        .Add "VT", "Transmisor (proximitor)": .Add "VS", "Switch de vibraci�n"
        .Add "ZE", "Sonda de posici�n": .Add "AY", "Convertidor IP"
        .Add "FY", "Convertidor IP": .Add "LY", "Convertidor IP"
        .Add "PY", "Convertidor IP": .Add "TY", "Convertidor IP"
        .Add "WE", "Celda de pesaje": .Add "WT", "Transmisor/Se�al Continua de peso"
        .Add "WI", "B�scula": .Add "CC", "Cup�n de corrosi�n"
        .Add "TMg", "Toma muestra": .Add "TML", "Toma muestra"
        .Add "XI", "Detector de paso Scrapper": .Add "IQ", "Inyecci�n de qu�mico"
End With

Dim dictDataSheets As Object
Set dictDataSheets = CreateObject("Scripting.Dictionary")
With dictDataSheets
    .Add "PI", "HD MANOMETROS"
    .Add "XI", "HD DETECTOR DE SCRAPER"
    .Add "PIT", "HD TRANSMISOR DE PRESION"
    .Add "PSV", "HD PSV"
    .Add "TI", "HD TERMOMETROS"
    .Add "TIT", "HD TRANSMISOR DE TEMPERATURA"
End With

Dim noSignalCodes As Variant, noSignalDict As Object, code As Variant
noSignalCodes = Array("AI", "FG", "LG", "AV", "FV", "HV", "LV", "PV", "PDV", "TV", "XV", _
                      "SDV", "BDV", "LBV", "LCV", "PCV", "PDCV", "TCV", "PSE", "TMg", "TML", "XI", "IQ")
Set noSignalDict = CreateObject("Scripting.Dictionary")
For Each code In noSignalCodes
    noSignalDict(code) = True
Next code

On Error Resume Next
Set acadApp = GetObject(, "AutoCAD.Application")
On Error GoTo 0

If acadApp Is Nothing Then
    MsgBox "No se detect� AutoCAD abierto. Por favor, abre AutoCAD y carga el archivo del cual se deben extraer los bloques.", vbCritical, "Error de conexi�n con AutoCAD"
    Exit Sub
End If

If acadApp.Documents.Count = 0 Then
    MsgBox "AutoCAD est� abierto pero no hay ning�n dibujo cargado. Abr� el plano antes de ejecutar la macro.", vbCritical, "Archivo no encontrado"
    Exit Sub
End If

Set acadDoc = acadApp.ActiveDocument
If acadDoc Is Nothing Then
    MsgBox "No se pudo acceder al documento activo de AutoCAD. Verific� que el dibujo est� correctamente cargado.", vbCritical, "Error con AutoCAD"
    Exit Sub
End If

Set modelSpace = acadDoc.modelSpace
If modelSpace Is Nothing Then
    MsgBox "No se pudo acceder al ModelSpace del archivo de AutoCAD. Puede que el archivo no sea v�lido o est� da�ado.", vbCritical, "Error con ModelSpace"
    Exit Sub
End If

hojaLI.Range("A7:C1000").ClearContents
filaExcel = 7

Dim foundCount As Long, foundRow As Long, term As String
Dim claveCompuesta As String
Dim i As Long

Application.StatusBar = "Extrayendo datos... 0% completado"
Dim bloquesProcesados As Long
bloquesProcesados = 0

For Each entidad In modelSpace
    If entidad.ObjectName = "AcDbBlockReference" Then
        Set bloqueRef = entidad
        If UCase(bloqueRef.EffectiveName) = "CO_INSTR" And bloqueRef.HasAttributes Then
            arrAttribs = bloqueRef.GetAttributes
            If IsArray(arrAttribs) Then
                valorFunction = "": valorTag = ""
                For i = LBound(arrAttribs) To UBound(arrAttribs)
                    Set atributo = arrAttribs(i)
                    Select Case UCase(atributo.TagString)
                        Case "FUNCTION": valorFunction = atributo.TextString
                        Case "TAG": valorTag = atributo.TextString
                    End Select
                Next i

                With hojaLI
                    .Cells(filaExcel, 1).Value = valorFunction
                    .Cells(filaExcel, 2).Value = valorTag
                    .Cells(filaExcel, 3).Value = IIf(typeMap.Exists(valorFunction), typeMap(valorFunction), "")
                    .Cells(filaExcel, 6).Value = codPID

                    claveCompuesta = valorFunction & "|" & valorTag
                    If valorFunction <> "" And valorTag <> "" Then
                        If tagTracking.Exists(claveCompuesta) Then
                            tagTracking(claveCompuesta) = tagTracking(claveCompuesta) & "," & filaExcel
                        Else
                            tagTracking.Add claveCompuesta, filaExcel
                        End If
                    End If

                    If noSignalDict.Exists(valorFunction) Then
                        .Range(.Cells(filaExcel, "J"), .Cells(filaExcel, "U")).Value = "-"
                    End If

                    If dictDataSheets.Exists(valorFunction) Then
                        term = dictDataSheets(valorFunction)
                        foundCount = 0
                        For iRowDocs = 3 To lastRowDocs
    If InStr(1, UCase(datosDocs(iRowDocs, 3)), term, vbTextCompare) > 0 Then
        foundCount = foundCount + 1
        foundRow = iRowDocs
    End If
Next iRowDocs
If foundCount = 1 Then
    .Cells(filaExcel, "G").Value = datosDocs(foundRow, 2)
Else
    .Cells(filaExcel, "G").Value = "-"
End If
                    Else
                        .Cells(filaExcel, "G").Value = "-"
                    End If
                End With

                filaExcel = filaExcel + 1
            End If
        End If
    End If

    bloquesProcesados = bloquesProcesados + 1
Application.StatusBar = "Extrayendo datos... " & Format(bloquesProcesados / modelSpace.Count, "0%") & " completado"
Next entidad

Application.StatusBar = False

' === Detección y marcado de tags duplicados ===
Dim tagKey As Variant
Dim filas() As String
Dim iAux As Long

For Each tagKey In tagTracking.Keys
    ' Si el tag aparece más de una vez (tiene más de una fila asociada)
    If InStr(tagTracking(tagKey), ",") > 0 Then
        cantidadDuplicados = cantidadDuplicados + 1
        filas = Split(tagTracking(tagKey), ",")
        For iAux = 0 To UBound(filas)
            ' Marca en rojo las columnas A y B de cada fila duplicada
            hojaLI.Cells(CLng(filas(iAux)), 1).Interior.Color = RGB(255, 199, 206)
            hojaLI.Cells(CLng(filas(iAux)), 2).Interior.Color = RGB(255, 199, 206)
        Next iAux
        ' Agrega el tag duplicado a la lista para el mensaje final
        listaDuplicados = listaDuplicados & "• " & tagKey & vbCrLf
    End If
Next

Dim totalExportados As Long
totalExportados = filaExcel - 7

If cantidadDuplicados > 0 Then
    MsgBox "Extracci�n completa: " & totalExportados & " instrumentos exportados." & vbCrLf & vbCrLf & _
           "Se detectaron " & cantidadDuplicados & " TAGs duplicados." & vbCrLf & _
           "Ver filas resaltadas en color rojo para m�s detalle.", vbExclamation, "Extracci�n finalizada con advertencias"
Else
    MsgBox "Extracci�n completa: " & totalExportados & " instrumentos exportados.", vbInformation, "Extracci�n finalizada"
End If

Call CompletarSenalesYUnidades

End Sub
    
' === RUTINA 2 ===
Sub CompletarSenalesYUnidades()
    Dim hojaLI As Worksheet
    Dim ultimaFila As Long, fila As Long
    Dim tipoInstrumento As String, tipoSenial As String
    Dim opcionesSenial As Variant, unidades As Variant
    Dim dictSeniales As Object, dictUnidades As Object

    Set hojaLI = ThisWorkbook.Sheets("LI")
    ultimaFila = hojaLI.Cells(hojaLI.Rows.Count, "A").End(xlUp).Row

    Set dictSeniales = CreateObject("Scripting.Dictionary")
    Set dictUnidades = CreateObject("Scripting.Dictionary")

    ' ============ TIPO DE SE�AL (columna K) ============
    dictSeniales.Add "PIT", Array("4-20mA", "4-20mA + HART")
    dictSeniales.Add "PDIT", Array("4-20mA", "4-20mA + HART")
    dictSeniales.Add "AIT", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "LIT", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "TIT", Array("4-20mA", "4-20mA + HART")
    dictSeniales.Add "TE", Array("mV", "Ohm")
    dictSeniales.Add "ST", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "VT", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "ZT", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "WT", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "BT", Array("4-20mA", "0-20mA", "1-5V")
    dictSeniales.Add "AT", Array("4-20mA", "0-20mA", "1-5V")

    dictSeniales.Add "PS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "PDS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "FS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "LS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "TS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "ZS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "ZSO", Array("Contacto seco", "24VDC")
    dictSeniales.Add "ZSC", Array("Contacto seco", "24VDC")
    dictSeniales.Add "BS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "SS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "VS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "XS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "YM", Array("Contacto seco", "24VDC")
    dictSeniales.Add "YR", Array("Contacto seco", "24VDC")
    dictSeniales.Add "YD", Array("Contacto seco", "24VDC")
    dictSeniales.Add "YS", Array("Contacto seco", "24VDC")
    dictSeniales.Add "YA", Array("Contacto seco", "24VDC")
    dictSeniales.Add "HS", Array("Contacto seco", "24VDC")

    dictSeniales.Add "AY", Array("4-20mA", "1-5V")
    dictSeniales.Add "FY", Array("4-20mA", "1-5V")
    dictSeniales.Add "LY", Array("4-20mA", "1-5V")
    dictSeniales.Add "PY", Array("4-20mA", "1-5V")
    dictSeniales.Add "TY", Array("4-20mA", "1-5V")
    dictSeniales.Add "XST", Array("4-20mA", "1-5V")
    dictSeniales.Add "XZT", Array("4-20mA", "1-5V")
    dictSeniales.Add "XET", Array("4-20mA", "1-5V")
    dictSeniales.Add "XIT", Array("4-20mA", "1-5V")
    dictSeniales.Add "XJT", Array("4-20mA", "1-5V")
    dictSeniales.Add "XGT", Array("4-20mA", "1-5V")
    dictSeniales.Add "XYT", Array("4-20mA", "1-5V")

    dictSeniales.Add "SV", Array("24VDC")
    dictSeniales.Add "XSMP", Array("24VDC")
    dictSeniales.Add "XSM", Array("24VDC")
    dictSeniales.Add "XSP", Array("24VDC")
    dictSeniales.Add "XSE", Array("24VDC")
    dictSeniales.Add "XSB", Array("24VDC")
    dictSeniales.Add "XSA", Array("24VDC")
    dictSeniales.Add "XSC", Array("24VDC")
    dictSeniales.Add "XSD", Array("24VDC")
    dictSeniales.Add "XY", Array("24VDC")

    ' ============ UNIDADES (columna U) ============
    dictUnidades.Add "PIT", Array("kg/cm2", "bar", "psi")
    dictUnidades.Add "TIT", Array("�C", "�F")
    dictUnidades.Add "LIT", Array("%", "m", "cm")
    dictUnidades.Add "FIT", Array("m3/h", "L/min", "gpm")
    dictUnidades.Add "WT", Array("kg", "ton")
    dictUnidades.Add "VT", Array("mm/s", "in/s")
    dictUnidades.Add "ST", Array("rpm", "Hz")

    For fila = 7 To ultimaFila
        tipoInstrumento = Trim(hojaLI.Cells(fila, "A").Value)
        If tipoInstrumento = "" Then GoTo Siguiente

        ' ========= Columna K =========
        If hojaLI.Cells(fila, "K").Value = "" Then
            Select Case tipoInstrumento
                Case "PIT", "PDIT", "AIT", "LIT", "TIT", "TE", "ST", "VT", "ZT", "WT", "BT", "AT"
                    tipoSenial = "AI"
                Case "PS", "PDS", "FS", "LS", "TS", "ZS", "ZSO", "ZSC", "BS", "SS", "VS", "XS", "YM", "YR", "YD", "YS", "YA", "HS"
                    tipoSenial = "DI"
                Case "AY", "FY", "LY", "PY", "TY", "XST", "XZT", "XET", "XIT", "XJT", "XGT", "XYT"
                    tipoSenial = "AO"
                Case "SV", "XSMP", "XSM", "XSP", "XSE", "XSB", "XSA", "XSC", "XSD", "XY"
                    tipoSenial = "DO"
                Case Else
                    tipoSenial = ""
            End Select
            If tipoSenial <> "" Then hojaLI.Cells(fila, "K").Value = tipoSenial
        End If

        ' ========= Columna L =========
        If dictSeniales.Exists(tipoInstrumento) Then
            hojaLI.Cells(fila, "L").Value = dictSeniales(tipoInstrumento)(0)
            With hojaLI.Range("L" & fila).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:=Join(dictSeniales(tipoInstrumento), ",")
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        Else
            hojaLI.Cells(fila, "L").Value = ""
        End If

        ' ========= Columna U =========
        If dictUnidades.Exists(tipoInstrumento) Then
            unidades = dictUnidades(tipoInstrumento)
            hojaLI.Cells(fila, "U").Value = unidades(0)
            If UBound(unidades) > 0 Then
                With hojaLI.Range("U" & fila).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=Join(unidades, ",")
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
            End If
        End If

        hojaLI.Range("L" & fila).WrapText = True
        hojaLI.Range("U" & fila).WrapText = True
        hojaLI.Rows(fila).AutoFit
Siguiente:
    Next fila

    MsgBox "Rutina de exportaci�n finalizada. Completar Servicio, Ubicaci�n e informaci�n de Alarmas, verificar y emitir", vbInformation
End Sub
