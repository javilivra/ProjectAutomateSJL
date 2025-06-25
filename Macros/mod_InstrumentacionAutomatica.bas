Attribute VB_Name = "Módulo1"
Option Explicit

Sub ExtraerAtributosBloqueInstrumentos()

Dim acadApp As Object
Dim acadDoc As Object
Dim modelSpace As Object
Dim entidad As Object
Dim bloqueRef As Object
Dim arrAttribs As Variant
Dim atributo As Object
Dim valorFunction As String
Dim valorTag As String
Dim filaExcel As Long
Dim tagTracking As Object
Set tagTracking = CreateObject("scripting.Dictionary")
Dim cantidadDuplicados As Long: cantidadDuplicados = 0
Dim listaDuplicados As String: listaDuplicados = ""

Dim wb As Workbook
Dim hojaLI As Worksheet
Dim hojaCar As Worksheet
Dim hojaNotasRef As Worksheet

' Variables para "LISTA DE DOCUMENTOS"
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

MsgBox "Por favor, espera mientras se extraen los datos desde AutoCAD. Este proceso puede tardar varios segundos según la cantidad de bloques."

matchRow = 0
For iRowDocs = 2 To lastRowDocs
    If InStr(1, UCase(hojaDocs.Cells(iRowDocs, "C").Value), "LISTA DE INSTRUMENTOS") > 0 Then
        matchRow = iRowDocs
        Exit For
    End If
Next iRowDocs
If matchRow = 0 Then
    wbDocs.Close SaveChanges:=False
    MsgBox "No se encontró 'LISTA DE INSTRUMENTOS' en la columna C.", vbExclamation
    Exit Sub
End If

codeAES = hojaDocs.Cells(matchRow, "A").Value
codeYPFProj = hojaDocs.Cells(matchRow, "B").Value
descDoc = hojaDocs.Cells(matchRow, "C").Value
codPID = hojaDocs.Cells(2, "B").Value

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
        .Add "AI", "Indicación de lectura analítica": .Add "AS", "Contacto o señal discreta"
        .Add "AX", "Contacto o señal discreta": .Add "AL", "Contacto o señal discreta"
        .Add "FE", "Elemento primario de Caudal": .Add "FG", "Visor en línea"
        .Add "FIT", "Caudalímetro": .Add "FI", "Indicación de Caudal"
        .Add "FS", "Contacto o señal discreta": .Add "FX", "Función/Cálculo"
        .Add "FF", "Relación de caudales": .Add "RO", "Orificio de restricción"
        .Add "FIC", "Indicador Controlador de Caudal": .Add "LE", "Elemento primario de medida de nivel"
        .Add "LG", "Nivel visual de vidrio": .Add "LI", "Indicación de Nivel"
        .Add "LIT", "Transmisor de Nivel": .Add "LS", "Contacto o señal discreta"
        .Add "LX", "Función/Cálculo": .Add "LIC", "Indicador Controlador de Nivel"
        .Add "PI", "Manómetro": .Add "PDI", "Manómetro Diferencial"
        .Add "PIT", "Transmisor de presión": .Add "PDIT", "Transmisor de presión Diferencial"
        .Add "PS", "Switch de Presión": .Add "PDS", "Switch de Presión Diferencial"
        .Add "PIC", "Indicador Controlador de Presión": .Add "TP", "Prueba de temperatura"
        .Add "TW", "Termovaina": .Add "TE", "Sensor de temperatura"
        .Add "TI", "Termómetro": .Add "TDI", "Termómetro diferencial"
        .Add "TIT", "Transmisor de temperatura": .Add "TS", "Contacto o señal discreta"
        .Add "TX", "Contacto o señal discreta": .Add "TL", "Contacto o señal discreta"
        .Add "TIC", "Indicador Controlador de Temperatura": .Add "AV", "Válvula de control"
        .Add "FV", "Válvula de control": .Add "HV", "Válvula de control"
        .Add "LV", "Válvula de control": .Add "PV", "Válvula de control"
        .Add "PDV", "Válvula de control": .Add "TV", "Válvula de control"
        .Add "XV", "Válvula ON-OFF": .Add "SDV", "Válvula Shutdown"
        .Add "BDV", "Válvula Blowdown": .Add "MOV", "Válvula Motorizada"
        .Add "LBV", "Válvula de corte de ductos": .Add "LCV", "Válvula autorreguladora por nivel"
        .Add "PCV", "Válvula autorreguladora por presión": .Add "PDCV", "Válvula autorreguladora por presión diferencial"
        .Add "TCV", "Válvula autorreguladora por temperatura": .Add "SV", "Válvula solenoide"
        .Add "ZS", "Switch de posición": .Add "ZT", "Transmisor de posición"
        .Add "ZSO", "Interruptor de posición Válvula abierta": .Add "ZSC", "Interruptor de posición Válvula cerrada"
        .Add "ZLO", "Indicación en pantalla de válvula abierta": .Add "ZLC", "Indicación en pantalla de válvula cerrada"
        .Add "PSE", "Disco de ruptura": .Add "PSV", "Válvula de seguridad/alivio"
        .Add "PVSV", "Válvula de presión/vacío": .Add "HS", "Pulsador"
        .Add "XL", "Lámpara": .Add "YL", "Lámpara"
        .Add "XA", "Alarma": .Add "XSMP", "Orden de marcha/paro"
        .Add "XSM", "Orden de marcha": .Add "XSP", "Orden de paro"
        .Add "XSE", "Permisivo de arranque": .Add "XSB", "Orden de disparo/señal de bloqueo"
        .Add "XSA", "Orden de abrir Válvula motorizada": .Add "XSC", "Orden de cerrar Válvula motorizada"
        .Add "XSD", "Orden de detener Válvula motorizada": .Add "XY", "-"
        .Add "YM", "Confirmación de marcha": .Add "YR", "Mando en remoto"
        .Add "YD", "Confirmación de equipo disponible": .Add "YS", "-"
        .Add "YA", "Estado de Falla": .Add "XST", "Consigna velocidad o frecuencia"
        .Add "XZT", "Consigna de posición": .Add "XET", "Consigna de tensión"
        .Add "XIT", "Consigna de intensidad": .Add "XJT", "Consigna de potencia"
        .Add "XGT", "Consigna de cos": .Add "XYT", "Otra variable a especificar"
        .Add "ST", "Velocidad o frecuencia": .Add "ET", "Tensión"
        .Add "IT", "Intensidad": .Add "JT", "Potencia"
        .Add "GT", "Cos": .Add "YT", "Otra variable a especificar"
        .Add "BE", "Detector de llama": .Add "BT", "Detector de llama"
        .Add "BI", "Indicación de llama": .Add "BS", "Contacto o señal discreta"
        .Add "BL", "Estado detector": .Add "SE", "Sonda de medida de velocidad"
        .Add "SS", "Contacto o señal discreta": .Add "VE", "Sonda de vibración"
        .Add "VT", "Transmisor (proximitor)": .Add "VS", "Switch de vibración"
        .Add "ZE", "Sonda de posición": .Add "AY", "Convertidor IP"
        .Add "FY", "Convertidor IP": .Add "LY", "Convertidor IP"
        .Add "PY", "Convertidor IP": .Add "TY", "Convertidor IP"
        .Add "WE", "Celda de pesaje": .Add "WT", "Transmisor/Señal Continua de peso"
        .Add "WI", "Báscula": .Add "CC", "Cupón de corrosión"
        .Add "TMg", "Toma muestra": .Add "TML", "Toma muestra"
        .Add "XI", "Detector de paso Scrapper": .Add "IQ", "Inyección de químico"
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
    MsgBox "No se detectó AutoCAD abierto. Por favor, abre AutoCAD y carga el archivo del cual se deben extraer los bloques.", vbCritical, "Error de conexión con AutoCAD"
    Exit Sub
End If

If acadApp.Documents.Count = 0 Then
    MsgBox "AutoCAD está abierto pero no hay ningún dibujo cargado. Abrí el plano antes de ejecutar la macro.", vbCritical, "Archivo no encontrado"
    Exit Sub
End If

Set acadDoc = acadApp.ActiveDocument
If acadDoc Is Nothing Then
    MsgBox "No se pudo acceder al documento activo de AutoCAD. Verificá que el dibujo esté correctamente cargado.", vbCritical, "Error con AutoCAD"
    Exit Sub
End If

Set modelSpace = acadDoc.modelSpace
If modelSpace Is Nothing Then
    MsgBox "No se pudo acceder al ModelSpace del archivo de AutoCAD. Puede que el archivo no sea válido o esté dañado.", vbCritical, "Error con ModelSpace"
    Exit Sub
End If

hojaLI.Range("A7:C1000").ClearContents
filaExcel = 7

Dim foundCount As Long, foundRow As Long, term As String
Dim claveCompuesta As String
Dim i As Long

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
                            If InStr(1, UCase(hojaDocs.Cells(iRowDocs, "C").Value), term, vbTextCompare) > 0 Then
                                foundCount = foundCount + 1
                                foundRow = iRowDocs
                            End If
                        Next iRowDocs
                        If foundCount = 1 Then
                            .Cells(filaExcel, "G").Value = hojaDocs.Cells(foundRow, "B").Value
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
Next entidad

wbDocs.Close SaveChanges:=False

Dim tagKey As Variant
Dim filas() As String
Dim iAux As Long

For Each tagKey In tagTracking.Keys
    If InStr(tagTracking(tagKey), ",") > 0 Then
        cantidadDuplicados = cantidadDuplicados + 1
        filas = Split(tagTracking(tagKey), ",")
        For iAux = 0 To UBound(filas)
            hojaLI.Cells(CLng(filas(iAux)), 1).Interior.Color = RGB(255, 199, 206)
            hojaLI.Cells(CLng(filas(iAux)), 2).Interior.Color = RGB(255, 199, 206)
        Next iAux
        listaDuplicados = listaDuplicados & "• " & tagKey & vbCrLf
    End If
Next

Dim totalExportados As Long
totalExportados = filaExcel - 7

If cantidadDuplicados > 0 Then
    MsgBox "Extracción completa: " & totalExportados & " instrumentos exportados." & vbCrLf & vbCrLf & _
           "Se detectaron " & cantidadDuplicados & " TAGs duplicados." & vbCrLf & _
           "Ver filas resaltadas en color rojo para más detalle.", vbExclamation, "Extracción finalizada con advertencias"
Else
    MsgBox "Extracción completa: " & totalExportados & " instrumentos exportados.", vbInformation, "Extracción finalizada"
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

    ' ============ TIPO DE SEÑAL (columna K) ============
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
    dictUnidades.Add "TIT", Array("°C", "°F")
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

    MsgBox "Rutina de exportación finalizada. Completar Servicio, Ubicación e información de Alarmas, verificar y emitir", vbInformation
End Sub
