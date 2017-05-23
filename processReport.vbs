'Author: Sergio Acosta Hernández
'Date 02/03/2012

'This script generates an HTML-formatted document with the current running processes in the system
'This report is saved in a custom location (please see 'carpeta' variable) and then a shortcut is created.


Option Explicit
'On Error Resume Next

Dim contenido, filasTabla, cabecera, titulo, cierre, escritorio, carpeta, estilos, nombreArchivo, ruta
Dim fso, wmi, shell, enlace, procesos, p, archivoTexto

'Folder to save the report will be "C:\informes"
carpeta = "C:\informes"
'Filename in spanish will be informe-yyyymmdd-hhMMss.html
nombreArchivo = Trim("informe-" & Year(Date()) & Month(Date()) & Day(Date()) & "-" & Replace(Time(), ":", "") & ".html")
'HTML and CSS code to format the table
cabecera = "<html>" & Chr(13) & "<meta name='generator' content='informe.vbs'>" & Chr(13) & "<title>Informe</title>" & Chr(13) & "<body>"
estilos = "<style type='text/css'>* {font-family: Arial, Helvetica, sans-serif;} #table-6 {width: 100%border: 1px solid #B0B0B0;}#table-6 tbody {margin: 0;padding: 0;border: 0;outline: 0;font-size: 100%;vertical-align: baseline;background: transparent;}#table-6 thead {text-align: left;}#table-6 thead th {background: -moz-linear-gradient(top, #F0F0F0 0, #DBDBDB 100%);background: -webkit-gradient(linear, left top, left bottom, color-stop(0%, #F0F0F0), color-stop(100%, #DBDBDB));filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#F0F0F0', endColorstr='#DBDBDB', GradientType=0);border: 1px solid #B0B0B0;color: #444;font-size: 16px;font-weight: bold;padding: 3px 10px;}#table-6 td {padding: 3px 10px;}#table-6 tr:nth-child(even) {background: #F2F2F2;}</style>"
titulo = "<h1>Informe procesos en ejecuci&oacute;n</h1>" & "<table id='table-6'> <tr> <thead> <th><b>Nombre del proceso</b></th> <th><b>PID</b></th> </thead> " 
cierre = "</body>" & Chr(13) & "</html>"

'Objetos necesarios para usar WMI y el sistema de archivos
Set wmi = GetObject("winmgmts:\\.\root\CIMV2")
Set fso = CreateObject("Scripting.FileSystemObject")
Set procesos = wmi.ExecQuery("select * from Win32_Process")

'Recorremos los procesos y los metemos en una tabla
For Each p In procesos
	filasTabla = filasTabla & Chr(13) & "<tr><td>" & LCase(p.Name) & "</td> <td>" & p.ProcessId & "</td> </tr>" & Chr(13)
Next

contenido = cabecera & titulo & estilos & filasTabla & cierre
ruta = carpeta & "\" & nombreArchivo

If  Not fso.FolderExists(carpeta) Then
'We do need to create the folder
	fso.CreateFolder (carpeta)
	Set archivoTexto = fso.OpenTextFile(ruta, 8, True)
	archivoTexto.WriteLine(contenido)
	archivoTexto.Close
Else
'We don't need to create the folder
	Set archivoTexto = fso.OpenTextFile(ruta, 8, True)
	archivoTexto.WriteLine(contenido)
	archivoTexto.Close
End If

Set shell = CreateObject("WScript.Shell")
escritorio = shell.SpecialFolders("Desktop")
'Generate link in desktop
Set enlace = shell.CreateShortcut(escritorio & "\Último Informe.lnk")

'Information for the shortcut
enlace.Description = "Informe de los procesos en ejecución"
enlace.HotKey = "CTRL+SHIFT+I"
enlace.IconLocation = "%SystemRoot%\system32\SHELL32.dll,1"
enlace.TargetPath = ruta
enlace.WorkingDirectory = carpeta
enlace.Save

'Spanish messagebox
Msgbox "Informe guardado en " & """" & carpeta & "\" & nombreArchivo & """. Se ha creado un acceso directo en el escritorio."


