' Wake on LAN multi-subred para soporte.
' Usa mc-wol.exe y un CSV de red para localizar MAC, IP, mascara y subred.
' Si origen y destino comparten subred, envia el WOL desde el equipo local.
' Si no la comparten, intenta usar un relay dentro de la subred remota.
' El CSV local puede refrescarse desde una ruta UNC de ejemplo antes del envio.
'==========
'Constantes
'==========
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001
Const OverwriteExisting = TRUE
'===========
'/Constantes
'===========

'=========
'Variables
'=========
'Destino=UCase (InputBox("Equipo a encender","Wake On Lan"))
'Destino=Ucase ("sconar007p")
Destino=UCase(Wscript.Arguments(0))

currentDirectory = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName))-(Len(WScript.ScriptName))) 'Directorio de ejecución del script

configPath = ResolveAbsolutePath(currentDirectory & "..\..\database\appsettings.json")
uncCSV = EnsureTrailingSlash(GetJsonSetting(configPath, "WolCsvShare", "\\server\share\network"))
nombreCSV = GetJsonSetting(configPath, "WolCsvFileName", "network_inventory_sample.csv")
CSVlocal=currentDirectory & nombreCSV
CSVremoto=uncCSV & nombreCSV

exeWOL="mc-wol.exe" 'http://www.matcode.com
destinoExeWOLRemoto="c$"
destinoExeWOLLocal="c:"
'==========
'/Variables
'==========

Function ResolveAbsolutePath(fPath)
	On Error Resume Next
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	ResolveAbsolutePath = objFSO.GetAbsolutePathName(fPath)
	If Err.Number <> 0 Then
		Err.Clear
		ResolveAbsolutePath = fPath
	End If
End Function

Function EnsureTrailingSlash(fPath)
	EnsureTrailingSlash = Trim(fPath)
	If EnsureTrailingSlash = "" Then Exit Function
	If Right(EnsureTrailingSlash, 1) <> "\" Then
		EnsureTrailingSlash = EnsureTrailingSlash & "\"
	End If
End Function

Function GetJsonSetting(fConfigPath, fKey, fDefaultValue)
	On Error Resume Next
	Dim objFSO, objFile, rawJson, objRegex, matches, value
	GetJsonSetting = fDefaultValue

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(fConfigPath) Then Exit Function

	Set objFile = objFSO.OpenTextFile(fConfigPath, 1)
	rawJson = objFile.ReadAll
	objFile.Close

	Set objRegex = CreateObject("VBScript.RegExp")
	objRegex.Global = False
	objRegex.IgnoreCase = True
	objRegex.Pattern = """" & fKey & """" & "\s*:\s*""([^""]*)"""

	If objRegex.Test(rawJson) Then
		Set matches = objRegex.Execute(rawJson)
		value = matches(0).SubMatches(0)
		value = Replace(value, "\\", "\")
		value = Replace(value, "\/", "/")
		GetJsonSetting = value
	End If
End Function

'=============================================
'Se comprueba el estado del equipo con un ping
'=============================================
Wscript.Echo ("************")
WScript.Echo ("PING DESTINO")
WScript.Echo ("************")
WScript.Echo ("Ping " & Destino)
rPing=Ping(Destino)
If (IsNull(rPing)) Then 
	WScript.Echo("Error en el ping.")
	WScript.Quit
ElseIf rPing=0 Then
	WScript.Echo ("Respuesta al ping.")
	WScript.Quit
Else
	WScript.Echo ("Sin respuesta al ping.")
'==============================================
'/Se comprueba el estado del equipo con un ping
'==============================================
	'===================
	'Se actualiza el CSV
	'===================
	Wscript.Echo ("*****************")
	WScript.Echo ("VERIFICACION CSV")
	WScript.Echo ("*****************")
	VerificaCSV CSVlocal,nombreCSV
	'===================
	'/Se actualiza el CSV
	'===================
	'==============================================
	'Se obtiene la DIRECCION MAC del equipo destino
	'==============================================
	WScript.Echo ("*************")
	WScript.Echo ("DIRECCION MAC")
	WScript.Echo ("*************")
	MAC=MACdestino(Destino,nombreCSV,currentDirectory)
	If IsNull (MAC) Then 'No se encuentra MAC
		WScript.Echo "No se encuentra DIRECCION MAC para el equipo: " & Destino
		WScript.Quit
	ElseIf MAC="" Then 'No se encuentra MAC
		WScript.Echo "No se encuentra DIRECCION MAC para el equipo: " & Destino
		WScript.Quit
	Else
		WScript.Echo "DIRECCION MAC del equipo " & Destino & ": " & MAC
	'===============================================
	'/Se obtiene la DIRECCION MAC del equipo destino
	'===============================================
		'========================
		'Se comparan las subredes
		'========================
		Wscript.Echo ("******")
		WScript.Echo ("SUBRED")
		WScript.Echo ("******")
		mSubLocal=SubRedLocal()
		mSubRemota=SubRedRemota(Destino,nombreCSV,currentDirectory)
		If mSubRemota="" OR IsNull(mSubRemota) Then
			WScript.Echo "CSV sin datos de subred remota."
			WScript.Echo "Obteniendo IP y máscara del equipo destino."
			WScript.Echo "Calculando subred remota."
			mSubRemota=IPMASK(Destino,nombreCSV,currentDirectory)
		End If
		If mSubLocal=mSubRemota Then
			WScript.Echo "Subred origen y subred destino iguales."
			WScript.Echo "Se envia WOL desde el equipo local."
			ejecucionLocal=EjecutaWOLLocal(exeWOL,CurrentDirectory,MAC) 'Ejecución local
			WScript.Echo EjecucionLocal
			VentanaPingFinal(Destino)
			WScript.Quit
		ElseIf mSubRemota="" OR IsNull(mSubRemota) Or mSubRemota="0.0.0.0" Then
			WScript.Echo "Sin datos de subred remota."
			WScript.Echo "Se envia WOL desde el equipo local."
			ejecucionLocal=EjecutaWOLLocal(exeWOL,CurrentDirectory,MAC) 'Ejecución local
			WScript.Echo EjecucionLocal
			VentanaPingFinal(Destino)
			WScript.Quit
		Else
			WScript.Echo "Subred origen y subred destino distintas."
			WScript.Echo "Se busca equipo relay."
		'=========================
		'/Se comparan las subredes
		'=========================
			'=====================
			'Se busca equipo relay
			'=====================
			Wscript.Echo ("*****")
			WScript.Echo ("RELAY")
			WScript.Echo ("*****")
			EquipoRelay=Relay(mSubRemota,nombreCSV,currentDirectory)
			If EquipoRelay="0" Then
				WScript.Echo "No se encuentra relay valido (sin equipos disponibles o sin permisos)."
				WScript.Echo "FALLBACK: Se envia WOL desde el equipo local."
				ejecucionLocal=EjecutaWOLLocal(exeWOL,CurrentDirectory,MAC) 'Ejecución local
				WScript.Echo EjecucionLocal
				VentanaPingFinal(Destino)
				WScript.Quit
			Else
				WScript.Echo "Se utilizara el equipo " & EquipoRelay & " para enviar el WOL."
				
				' Obtener IP y Máscara del destino para calcular broadcast
				IPDestino = ObtenerIP(Destino, nombreCSV, currentDirectory)
				MaskDestino = ObtenerMask(Destino, nombreCSV, currentDirectory)
				BroadcastDestino = CalcBroadcastAddress(IPDestino, MaskDestino)
				WScript.Echo "IP destino: " & IPDestino & " / Broadcast: " & BroadcastDestino
			'======================
			'/Se busca equipo relay
			'======================
				'================
				'Ejecución remota
				'================
				WScript.Echo "Copiando " & ExeWOL & " en \\" & EquipoRelay & "\" & destinoExeWOLRemoto
				Copia=CopiaWOL(EquipoRelay,ExeWOL,destinoExeWOLRemoto,currentDirectory)
				WScript.Echo Copia
				
				' Verificar si la copia fue exitosa
				If InStr(Copia, "[ERROR]") > 0 Then
					WScript.Echo "[ERROR] No se pudo copiar el archivo al relay."
					WScript.Echo "FALLBACK: Se envia WOL desde el equipo local."
					ejecucionLocal=EjecutaWOLLocal(exeWOL,CurrentDirectory,MAC)
					WScript.Echo EjecucionLocal
					VentanaPingFinal(Destino)
					WScript.Quit
				End If
				
				WScript.Echo "Ejecutando WOL desde " & EquipoRelay
				Ejecucion=EjecutaWOLRemoto(EquipoRelay,exeWOL,DestinoExeWOLLocal,MAC,BroadcastDestino,currentDirectory) 'Ejecución remota
				WScript.Echo Ejecucion
				
				' Verificar si la ejecución fue exitosa
				If InStr(Ejecucion, "Error") > 0 Or InStr(Ejecucion, "[ERROR]") > 0 Then
					WScript.Echo "[ERROR] Fallo en la ejecucion remota."
					WScript.Echo "FALLBACK: Se envia WOL desde el equipo local."
					ejecucionLocal=EjecutaWOLLocal(exeWOL,CurrentDirectory,MAC)
					WScript.Echo EjecucionLocal
					VentanaPingFinal(Destino)
					WScript.Quit
				End If
				
				WScript.Echo ""
				WScript.Echo "Esperando que el equipo encienda..."
				WScript.Sleep (5000) ' Esperar 5 segundos para que el equipo inicie
				WScript.Echo "Eliminando " & "\\" & EquipoRelay & "\" & destinoExeWOLRemoto & "\" & ExeWOL
				Elimina=EliminaWOL(EquipoRelay,ExeWOL,destinoExeWOLRemoto)
				If Elimina <> "" Then
					WScript.Echo Elimina
				End If
				VentanaPingFinal(Destino)
				'=================
				'/Ejecución remota
				'=================
			End If	
		End If
	End If
End If

Sub VerificaCSV(fCSVlocal,fNombreCSV)
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Verificar si el CSV local existe
	WScript.Echo "Verificando " & fNombreCSV & " local..."
	
	If Not objFSO.FileExists(fCSVlocal) Then
		' No existe local - intentar copiar del servidor
		WScript.Echo "[AVISO] " & fNombreCSV & " no encontrado localmente."
		WScript.Echo "Intentando descargar desde servidor..."
		
		If objFSO.FileExists(CSVremoto) Then
			objFSO.CopyFile CSVremoto, fCSVlocal, True
			If Err.Number = 0 Then
				WScript.Echo "[OK] " & fNombreCSV & " descargado desde servidor."
			Else
				WScript.Echo "[ERROR] No se pudo copiar el CSV: " & Err.Description
				WScript.Echo "Verifica permisos de red."
				WScript.Quit
			End If
		Else
			WScript.Echo "[ERROR] No se encuentra " & fNombreCSV & " en el servidor: " & CSVremoto
			WScript.Quit
		End If
	Else
		' Existe local - comparar fechas con el servidor
		WScript.Echo "[OK] " & fNombreCSV & " encontrado localmente."
		
		If objFSO.FileExists(CSVremoto) Then
			Set fileLocal = objFSO.GetFile(fCSVlocal)
			Set fileRemoto = objFSO.GetFile(CSVremoto)
			
			If fileRemoto.DateLastModified > fileLocal.DateLastModified Then
				WScript.Echo "Actualizando CSV desde servidor (version mas reciente disponible)..."
				objFSO.CopyFile CSVremoto, fCSVlocal, True
				If Err.Number = 0 Then
					WScript.Echo "[OK] CSV actualizado correctamente."
				Else
					WScript.Echo "[AVISO] No se pudo actualizar: " & Err.Description
					WScript.Echo "Usando version local existente."
				End If
			Else
				WScript.Echo "[OK] CSV local esta actualizado."
			End If
		Else
			WScript.Echo "[AVISO] No se puede acceder al servidor. Usando CSV local."
		End If
	End If
End Sub

Function MACdestino(fDestino,fNombreCSV,fCurrentDirectory)
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCurrentDirectory & fNombreCSV, 1)
	
	MACdestino = ""
	Dim lineNum : lineNum = 0
	
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		lineNum = lineNum + 1
		
		If lineNum = 1 Then
			' Saltar encabezado
		Else
			' Parsear CSV: RESOURCE_GUID,NAME,MAC,IP,MASK,SUBNET
			arrFields = Split(strLine, ",")
			If UBound(arrFields) >= 5 Then
				strName = Trim(arrFields(1))
				If UCase(strName) = UCase(fDestino) Then
					MACdestino = Trim(arrFields(2))
					WScript.Echo "[DEBUG] MAC encontrada: " & MACdestino
					Exit Do
				End If
			End If
		End If
	Loop
	
	objFile.Close
	
	If MACdestino = "" Then
		WScript.Echo "[ERROR] No se encontro registro para: " & fDestino
	End If
End Function

Function Relay(fSubNet,fnombreCSV,fCurrentDirectory)
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCurrentDirectory & fnombreCSV, 1)
	
	Relay = "0"
	Dim lineNum : lineNum = 0
	Dim arrCandidatos() : Dim candidateCount : candidateCount = 0
	fTempsubnet = Replace(fSubNet, ".", ":")
	
	' Primera pasada: recopilar todos los candidatos de la subred
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		lineNum = lineNum + 1
		
		If lineNum = 1 Then
			' Saltar encabezado
		Else
			' Parsear CSV: RESOURCE_GUID,NAME,MAC,IP,MASK,SUBNET
			arrFields = Split(strLine, ",")
			If UBound(arrFields) >= 5 Then
				strSubnet = Trim(arrFields(5))
				strName = UCase(Trim(arrFields(1)))
				
				' Verificar si está en la subred correcta y NO empieza por "H"
				If strSubnet = fTempsubnet And Left(strName, 1) <> "H" Then
					' Agregar a la lista de candidatos
					ReDim Preserve arrCandidatos(candidateCount)
					arrCandidatos(candidateCount) = strName
					candidateCount = candidateCount + 1
				End If
			End If
		End If
	Loop
	objFile.Close
	
	' Segunda pasada: probar cada candidato
	If candidateCount > 0 Then
		WScript.Echo "Equipos candidatos encontrados en la subred (excluyendo H*): " & candidateCount
		
		For i = 0 To UBound(arrCandidatos)
			strCandidato = arrCandidatos(i)
			WScript.Echo "Probando relay: " & strCandidato
			
			' Verificar ping
			If Ping(strCandidato) = 0 Then
				WScript.Echo "  [OK] Responde al ping"
				
				' Verificar permisos intentando acceder a C$
				If TestRemoteAccess(strCandidato) Then
					WScript.Echo "  [OK] Permisos verificados"
					Relay = strCandidato
					Exit Function
				Else
					WScript.Echo "  [AVISO] Sin permisos administrativos, buscando otro..."
				End If
			Else
				WScript.Echo "  [AVISO] No responde al ping"
			End If
		Next
		
		WScript.Echo "[AVISO] No se encontró ningún relay con permisos válidos"
	Else
		WScript.Echo "[AVISO] No se encontraron equipos candidatos en la subred (se excluyen equipos H*)"
	End If
	
	' Si llegamos aquí, no se encontró relay válido
	Relay = "0"
End Function

Function TestRemoteAccess(fComputerName)
	' Verifica si tenemos permisos administrativos intentando listar C$
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	strRemotePath = "\\" & fComputerName & "\C$"
	
	' Intentar acceder a la carpeta remota
	If objFSO.FolderExists(strRemotePath) Then
		' Intentar listar contenido para verificar permisos reales
		Set objFolder = objFSO.GetFolder(strRemotePath)
		Set objFiles = objFolder.Files
		
		' Si llegamos aquí sin error, tenemos permisos
		TestRemoteAccess = True
	Else
		TestRemoteAccess = False
	End If
	
	' Si hubo error, no tenemos permisos
	If Err.Number <> 0 Then
		TestRemoteAccess = False
		Err.Clear
	End If
End Function

Function Ping(fHost)
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & fHost & "'")
	For Each objStatus in objPing
		Ping=objStatus.StatusCode
	Next
End Function

Function CopiaWOL(fEquipoRelay,fExeWOL,fDestinoExeWOLRemoto,fCurrentDirectory)
	On Error Resume Next
	fOrigenExeWOL=fCurrentDirectory + fExeWOL
	fDestinoExeWOL="\\" & fEquipoRelay & "\" & fDestinoExeWOLRemoto & "\"
	fArchivoCompleto = fDestinoExeWOL & fExeWOL
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Verificar si el archivo ya existe en el relay
	If objFSO.FileExists(fArchivoCompleto) Then
		CopiaWOL = "[OK] " & fExeWOL & " ya existe en " & fEquipoRelay
	Else
		' Intentar copiar el archivo
		objFSO.CopyFile fOrigenExeWOL, fDestinoExeWOL, OverwriteExisting
		If Err.Number <> 0 Then
			CopiaWOL = "[ERROR] No se pudo copiar " & fExeWOL & ": " & Err.Description
		Else
			CopiaWOL = "[OK] " & fExeWOL & " copiado correctamente"
		End If
	End If
End Function

Function EliminaWOL(fEquipoRelay,fExeWOL,fDestinoExeWOLRemoto)
	On Error Resume Next
	RutaExeWOL="\\" & fEquipoRelay & "\" & fDestinoExeWOLRemoto & "\" & fExeWOL
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile(RutaExeWOL)
	If Err.Number <> 0 Then
		EliminaWOL = "[AVISO] No se pudo eliminar " & fExeWOL & " del relay (puede estar en uso). Error: " & Err.Description
	Else
		EliminaWOL = fExeWOL & " eliminado correctamente."
	End If
End Function

Function SubRedRemota(fDestino,fCSV,frutaCSV)
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(frutaCSV & fCSV, 1)
	
	SubRedRemota = ""
	Dim lineNum : lineNum = 0
	
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		lineNum = lineNum + 1
		
		If lineNum = 1 Then
			' Saltar encabezado
		Else
			' Parsear CSV: RESOURCE_GUID,NAME,MAC,IP,MASK,SUBNET
			arrFields = Split(strLine, ",")
			If UBound(arrFields) >= 5 Then
				strName = Trim(arrFields(1))
				If UCase(strName) = UCase(fDestino) Then
					fSubRed = Trim(arrFields(5))
					SubRedRemota = Replace(fSubRed, ":", ".")
					Exit Do
				End If
			End If
		End If
	Loop
	
	objFile.Close
End Function

Function IPMASK(fDestino,fnombreCSV,fCurrentDirectory) 'Obtiene IP y Máscara del equipo remoto y calcula la subred
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCurrentDirectory & fnombreCSV, 1)
	
	IPMASK = ""
	Dim lineNum : lineNum = 0
	
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		lineNum = lineNum + 1
		
		If lineNum = 1 Then
			' Saltar encabezado
		Else
			' Parsear CSV: RESOURCE_GUID,NAME,MAC,IP,MASK,SUBNET
			arrFields = Split(strLine, ",")
			If UBound(arrFields) >= 5 Then
				strName = Trim(arrFields(1))
				If UCase(strName) = UCase(fDestino) Then
					fIP = Replace(Trim(arrFields(3)), ":", ".")
					fMask = Replace(Trim(arrFields(4)), ":", ".")
					IPMASK = CalcNetworkAddress(fIP, fMask)
					Exit Do
				End If
			End If
		End If
	Loop
	
	objFile.Close
End Function

Function SubRedLocal() 'Obtiene la subred a partir de la IP y la máscara
	fstrComputer="."
	Set objWMIService = GetObject("winmgmts:\\" & fstrComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each objItem in colItems
		If InStr(1,objItem.Caption,"virtual",1)=0 Then
			fIP = objItem.IPAddress(0)
			fMask=objItem.IPSubnet(0)
		End If
	Next
	SubRedLocal=CalcNetworkAddress(fIP,fMask)
End Function


Function EjecutaWOLLocal(fexeWOL,fCurrentDirectory,fMAC)
	' mc-wol.exe usa formato xx:xx:xx:xx:xx:xx (NO cambiar los dos puntos)
	' Construir ruta completa con comillas para manejar espacios
	ComandoWOL= """" & fCurrentDirectory & fexeWOL & """ " & fMAC
	WScript.Echo "[DEBUG] Comando WOL local: " & ComandoWOL
	Set wshShell = WScript.CreateObject ("WSCript.shell")
	EjecutaWOLLocal=wshShell.Run (ComandoWOL,8,TRUE)
End Function

Function EjecutaWOLRemoto(fEquipoRelay,fexeWOL,fDestinoExeWOLLocal,fMAC,fBroadcast,fCurrentDirectory)
	' mc-wol.exe usa formato xx:xx:xx:xx:xx:xx (NO cambiar los dos puntos)
	' Agregar parámetro /a con DIRECCION de broadcast
	ComandoWOL= fDestinoExeWOLLocal & "\" & fexeWOL & " " & fMAC & " /a " & fBroadcast
	WScript.Echo "[DEBUG] Comando WOL remoto: " & ComandoWOL
	
	' Usar PsExec para ejecutar remotamente (más confiable que WMI)
	ComandoPsExec = """" & fCurrentDirectory & "PsExec.exe"" -accepteula \\" & fEquipoRelay & " " & ComandoWOL
	WScript.Echo "[DEBUG] Ejecutando via PsExec: " & ComandoPsExec
	
	Set wshShell = WScript.CreateObject ("WSCript.shell")
	returnCode = wshShell.Run (ComandoPsExec, 0, TRUE)
	
	If returnCode = 0 Then
		EjecutaWOLRemoto= "Magic Packet enviado desde " & fEquipoRelay & " [OK]"
	Else
		EjecutaWOLRemoto= "Error en ejecucion remota. Codigo: " & returnCode
	End If
End Function

Function VentanaPingFinal(fDestino)
	Set wshShell = WScript.CreateObject ("WSCript.shell")
	t=wshShell.Run ("CMD /C PING -t " & fDestino,1,FALSE)
End Function

Function ObtenerIP(fDestino, fnombreCSV, fCurrentDirectory)
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCurrentDirectory & fnombreCSV, 1)
	
	ObtenerIP = ""
	Dim lineNum : lineNum = 0
	
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		lineNum = lineNum + 1
		
		If lineNum > 1 Then
			arrFields = Split(strLine, ",")
			If UBound(arrFields) >= 5 Then
				strName = Trim(arrFields(1))
				If UCase(strName) = UCase(fDestino) Then
					ObtenerIP = Replace(Trim(arrFields(3)), ":", ".")
					Exit Do
				End If
			End If
		End If
	Loop
	objFile.Close
End Function

Function ObtenerMask(fDestino, fnombreCSV, fCurrentDirectory)
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(fCurrentDirectory & fnombreCSV, 1)
	
	ObtenerMask = ""
	Dim lineNum : lineNum = 0
	
	Do Until objFile.AtEndOfStream
		strLine = objFile.ReadLine
		lineNum = lineNum + 1
		
		If lineNum > 1 Then
			arrFields = Split(strLine, ",")
			If UBound(arrFields) >= 5 Then
				strName = Trim(arrFields(1))
				If UCase(strName) = UCase(fDestino) Then
					ObtenerMask = Replace(Trim(arrFields(4)), ":", ".")
					Exit Do
				End If
			End If
		End If
	Loop
	objFile.Close
End Function

Function CalcBroadcastAddress(strIP, strMask)
	' Calcula la DIRECCION de broadcast: IP OR (NOT Mask)
	Dim strBinIP : strBinIP = ConvertIPToBinary(strIP)
	Dim strBinMask : strBinMask = ConvertIPToBinary(strMask)
	
	' Bitwise OR con NOT de la máscara
	Dim i, strBinBroadcast
	For i = 1 to Len(strBinIP)
		Dim strIPBit : strIPBit = Mid(strBinIP, i, 1)
		Dim strMaskBit : strMaskBit = Mid(strBinMask, i, 1)
		
		If strIPBit = "." Then
			strBinBroadcast = strBinBroadcast & strIPBit
		ElseIf strMaskBit = "0" Then
			strBinBroadcast = strBinBroadcast & "1"
		Else
			strBinBroadcast = strBinBroadcast & strIPBit
		End If
	Next
	
	CalcBroadcastAddress = ConvertBinIPToDecimal(strBinBroadcast)
End Function

'=============================COPIADO================================
'http://www.indented.co.uk/index.php/2008/10/21/vbscript-subnet-math/

Function CalcNetworkAddress(strIP, strMask)
  ' Generates the Network Address from the IP and Mask
  
  ' Conversion of IP and Mask to binary
  Dim strBinIP : strBinIP = ConvertIPToBinary(strIP)
  Dim strBinMask : strBinMask = ConvertIPToBinary(strMask)
 
  ' Bitwise AND operation (except for the dot)
  Dim i, strBinNetwork
  For i = 1 to Len(strBinIP)
    Dim strIPBit : strIPBit = Mid(strBinIP, i, 1)
    Dim strMaskBit : strMaskBit = Mid(strBinMask, i, 1)
 
    If strIPBit = "1" And strMaskBit = "1" Then
      strBinNetwork = strBinNetwork & "1"
    ElseIf strIPBit = "." Then
      strBinNetwork = strBinNetwork & strIPBit
    Else
      strBinNetwork = strBinNetwork & "0"
    End If
  Next
 
  ' Conversion of Binary IP to Decimal
  CalcNetworkAddress= ConvertBinIPToDecimal(strBinNetwork)
End Function

Function ConvertBinIPToDecimal(strBinIP)
  ' Convert binary form of an IP back to decimal
 
  Dim arrOctets : arrOctets = Split(strBinIP, ".")
  Dim i
  For i = 0 to UBound(arrOctets)
    Dim intOctet : intOctet = 0
    Dim j
    For j = 0 to 7
      Dim intBit : intBit = CInt(Mid(arrOctets(i), j + 1, 1))
      If intBit = 1 Then
        intOctet = intOctet + 2^(7 - j)
      End If
    Next
    arrOctets(i) = CStr(intOctet)
  Next
 
  ConvertBinIPToDecimal = Join(arrOctets, ".")
End Function

Function ConvertIPToBinary(strIP)
  ' Converts an IP Address into Binary
 
  Dim arrOctets : arrOctets = Split(strIP, ".")
  Dim i
  For i = 0 to UBound(arrOctets)
    Dim intOctet : intOctet = CInt(arrOctets(i))
    Dim strBinOctet : strBinOctet = ""
    Dim j
    For j = 0 To 7
      If intOctet And (2^(7 - j)) Then
        strBinOctet = strBinOctet & "1"
      Else
        strBinOctet = strBinOctet & "0"
      End If
    Next
    arrOctets(i) = strBinOctet
  Next
  ConvertIPToBinary = Join(arrOctets, ".")
End Function

'http://www.indented.co.uk/index.php/2008/10/21/vbscript-subnet-math/
'=============================COPIADO================================
