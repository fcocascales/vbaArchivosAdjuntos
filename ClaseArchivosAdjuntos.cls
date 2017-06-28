VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClaseArchivosAdjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'==================================================
' ClaseArchivosAdjuntos
'
' Este código añade la posibilidad de adjuntar ficheros a una tabla desde un formulario
' Los ficheros se guardan en subcarpetas ubicadas a partir de dónde se encuentre el archivo mdb
'
' Módulo original de "http://siliconproject.com.ar/neckkito/"
'
' Refactorizado por Francisco Cascales - 2017-03-26
'
'    - Flexibilidad en el nombre de los campos
'    - Facilitar al máximo su uso desde el formulario
'    - Poder ser usado en formularios no independientes
'    - Renombrar variables
'    - Creación de nuevas funciones para no repetir código
'    - Reescribir comentarios y eliminar otros
'    - Eliminar variables globales usadas como locales
'    - Facilitar renombrado de las carpetas
'    - Reescritura de código para facilitar su compresión
'    - Convertirlo en un módulo de Clase
'    - Añadir propiedad CarpetaDatos para separar datos por tablas
'    - Añadir propiedad NombreTabla para indicar la ruta raíz a la BD de una tabla vinculada
'    - Añadir propiedad CarpetaRaiz para ubicar exactamente los adjuntos
'
'
'==================================================
' REQUISITOS
'
' Registrar estas librerías (>Herramientas>Referencias)
'
' -> Visual Basic For Applications
' -> Microsoft Access xx.x Object Library
' -> Microsoft Office xx.x Access database engine Object Library // Microsoft DAO Object Library
' -> Microsoft Office xx.x Object Library
' -> Microsoft Scripting Runtime
'
'==================================================
' FORMA DE EMPLEO
'
' TABLA
'
' En la tabla se necesita que exista:
'
'   - un campo único [id]
'   - y un campo de tipo memo o texto largo [adjuntos]
'
' FORMULARIO
'
' Se necesitan 3 controles:
'
'  1) Un id que debe ser único y es como se llamarán las subcarpetas dentro de Adjuntos/Datos
'  2) Un campo para guardar el nombre de los adjuntos. Debe ser de tipo texto. Debería ser visible=no
'  3) Un combobox independiente para seleccionar el documento adjunto actual
'
' Variables del formulario
'   private archivosAdjuntos as ClasearchivosAdjuntos
'
' Eventos de formulario:
'
'   Form_Load=>
'       Set archivosAdjuntos = New ClaseArchivosAdjuntos
'       archivosAdjuntos.NombreTabla = "Datos"
'       archivosAdjuntos.CampoId = me.id
'       archivosAdjuntos.CampoAdjuntos = me.adjuntos
'       archivosAdjuntos.ComboboxIndependiente = me.comboDocumento
'       archivosAdjuntos.Inicializar
'
'   Form_AfterInsert=>
'       Call archivosAdjuntos.Procesar
'   Form_AfterUpdate=>
'       Call archivosAdjuntos.Procesar
'   Form_Current=>
'       Call archivosAdjuntos.Actualizar
'       Call archivosAdjuntos.Asear
'
' Botones:
'
'   Boton1=>
'       Call archivosAdjuntos.BotonElegir
'   Boton2=>
'       Call archivosAdjuntos.BotonVer
'   Boton3=>
'       Call archivosAdjuntos.BotonExportar
'   Boton4=>
'       Call archivosAdjuntos.BotonBorrar
'
'==================================================

Private Const CARPETA_RAIZ = "Adjuntos"
Private Const CARPETA_TEMPORAL = "Temp"
Private Const CARPETA_DATOS = "Datos"

Private strCarpetaDatos As String
Private strNombreTabla As String
Private strCarpetaRaiz As String
Private textboxId As Textbox
Private textboxAdjuntos As Textbox
Private comboDocumento As Combobox

Private Sub Class_Initialize()
    strCarpetaDatos = CARPETA_DATOS
    strNombreTabla = ""
    strCarpetaRaiz = ""
    Set textboxId = Nothing
    Set textboxAdjuntos = Nothing
    Set comboDocumento = Nothing
End Sub

Public Property Let CampoId(object As Textbox)
    Set textboxId = object
End Property
Public Property Let CampoAdjuntos(object As Textbox)
    Set textboxAdjuntos = object
End Property
Public Property Let ComboboxIndependiente(object As Combobox)
    Set comboDocumento = object
End Property
Public Property Let NombreTabla(value As String)
    strNombreTabla = value
    strCarpetaDatos = value
    strCarpetaRaiz = obtenerCarpetaTabla()
End Property
Public Property Let CarpetaDatos(value As String)
    strCarpetaDatos = value
End Property
Public Property Let CarpetaRaiz(value As String)
    strCarpetaRaiz = value
End Property

' Usado después de ajustar las propiedades de arriba
'
Public Sub Inicializar()
    Call crearCarpetas
    Call vaciarTemporal
End Sub

' Usado en el evento Current del formulario
' Actualiza el origen de la lista del combo según el campo de adjuntos de la tabla
'
Public Sub Actualizar()
    Dim adjuntos As String
    Dim longitud As Integer
    
    adjuntos = Nz(textboxAdjuntos, "")
    comboDocumento.RowSource = adjuntos
    comboDocumento.Requery
    longitud = comboDocumento.ListCount
    If longitud = 0 Then
        comboDocumento = ""
    Else
        comboDocumento = comboDocumento.ItemData(longitud - 1)
    End If
End Sub

' Usado en el primer botón del formulario para ELEGIR el archivo adjunto
' Selecciona un documento, lo carga en el combo del formulario y lo guarda en la carpeta Temp
'
Public Sub BotonElegir()
    On Error GoTo Errores
    
    Dim fso As Scripting.FileSystemObject
    Dim arch As Scripting.File
    Dim myTemp As String
    Dim myFile As String
    
    myTemp = getRutaTemporal()
    myFile = dialogoSeleccionarArchivo()
    If IsNull(myFile) Or myFile = "" Then Exit Sub
    
    'Lo copiamos a la carpeta Temp
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arch = fso.GetFile(myFile)
    arch.Copy (agregarRuta(myTemp, arch.Name))
    
    'Lo añadimos al cuadro de texto oculto
    If IsNull(textboxAdjuntos) Then
        textboxAdjuntos = arch.Name & ";"
    Else
        textboxAdjuntos = textboxAdjuntos & arch.Name & ";"
    End If

Salida:
    'Lo añadimos al combo
    Call Actualizar
    Exit Sub
Errores:
    Select Case Err.Number
        Case 53
            'MsgBox "El archivo """ & miArchivo & """ no existe.", vbInformation + vbOKOnly, "ERROR"
            MsgBox "El archivo no existe.", vbInformation + vbOKOnly, "ERROR"
        Case Else
            Call mostrarError("ElegirAdjuntos")
    End Select
    Resume Salida
End Sub


' Usado en el segundo botón del formulario para ver el adjunto seleccionado en el combobox
'
Public Sub BotonVer()
    Dim myFile As String
    Dim myPath As String
    
    myFile = Nz(comboDocumento, "")
    If myFile = "" Then Exit Sub
        
    myPath = buscarRutaArchivo(textboxId, myFile)
    If myPath = "" Then Exit Sub
        
    Application.FollowHyperlink agregarRuta(myPath, myFile)
End Sub

' Usado en el tercer botón del formulario para EXPORTAR el adjunto seleccionado en el combobox
'
Public Sub BotonExportar()
    'on error goto errores

    Dim myFile As String
    Dim myPath As String
    Dim myTarget As String
    Dim fso As Object
    Dim arch As Object
    
    myFile = Nz(comboDocumento, "")
    If myFile = "" Then Exit Sub
    
    myPath = buscarRutaArchivo(textboxId, myFile)
    If myPath = "" Then Exit Sub
        
    myTarget = dialogoSeleccionarCarpeta()
    If myTarget = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arch = fso.GetFile(agregarRuta(myPath, myFile))
    arch.Copy (agregarRuta(myTarget, arch.Name))
End Sub

' Usado en el cuarto botón del formulario para BORRAR el adjunto seleccionado en el combobox
'
Public Sub BotonBorrar()
    Dim resp As Integer
    Dim myFile As String
    Dim i As Integer
        
    myFile = Nz(comboDocumento, "")
    If myFile = "" Then Exit Sub
    resp = MsgBox("¿Está seguro que desea eliminar el archivo """ & myFile & """?", vbInformation + vbYesNo, "CONFIRMA BORRADO")
    If resp = vbNo Then Exit Sub
    
    textboxAdjuntos = ""
    For i = 0 To comboDocumento.ListCount
        If comboDocumento.ItemData(i) <> myFile Then
            textboxAdjuntos = textboxAdjuntos & comboDocumento.ItemData(i) & ";"
        End If
    Next
    Call Actualizar
End Sub


' Usado después de guardar; en los eventos AfterInsert y AfterUpdate del formulario
' Pasar de la carpeta temporal a la definitiva los adjuntos
'
Public Sub Procesar()
On Error GoTo Errores
    Dim fso As Object
    Dim fsc As Object
    Dim fsa As Object
    Dim fsas As Object
    Dim myTemp As String
    Dim myData As String
    
    'Establecemos la ruta en función del formulario
    myTemp = getRutaTemporal()
    myData = getRutaTabla()
    
    'Movemos los docus de Temp a su destino
    MkDir myData
    MkDir agregarRuta(myData, textboxId)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsc = fso.GetFolder(myTemp)
    Set fsas = fsc.Files
    'Recorremos los archivos de la carpeta
    For Each fsa In fsas
        fsa.Move (agregarRuta(myData, textboxId, fsa.Name))
    Next fsa
    
Salida:
    Set fsa = Nothing
    Set fsas = Nothing
    Set fsc = Nothing
    Set fso = Nothing
    Exit Sub
Errores:
    Select Case Err.Number
        Case 53 'el archivo no existe
            Resume Next
        Case 75, 76 'la carpeta ya existe
            Resume Next
        Case Else
            Call mostrarError("ProcesarAdjuntos")
    End Select
End Sub


' Usado en el evento Current del formulario
' Borra el archivo adjunto si no está en el combo
'
Public Sub Asear()
On Error Resume Next
    Dim fso As Object
    Dim fsc As Object
    Dim fsa As Object
    Dim fsas As Object
    Dim existeDoc As Boolean
    Dim myId As String
    Dim i As Integer
     
    myId = getRutaId(textboxId)

    If Dir(myId, vbDirectory) = "" Then GoTo Salida 'Comprobamos que existe la carpeta
    
    If comboDocumento.ListCount = 0 Then
        Kill agregarRuta(myId, "*.*")
        RmDir myId
        GoTo Salida
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsc = fso.GetFolder(myId)
    Set fsas = fsc.Files
    'Recorremos los archivos de la carpeta
    For Each fsa In fsas
        existeDoc = False
        For i = 0 To comboDocumento.ListCount
            If comboDocumento.ItemData(i) = fsa.Name Then
                existeDoc = True
            End If
        Next i
        If existeDoc = False Then Kill agregarRuta(myId, fsa.Name)
    Next fsa
    
Salida:
    Set fsas = Nothing
    Set fsc = Nothing
    Set fso = Nothing
End Sub


'==================================================
' PRIVADO

' Usado en el evento Load del formulario
' Creamos, si no existen, las carpetas "Documentos" y "Temp"
'
Private Sub crearCarpetas()
    On Error Resume Next
    
    Dim ruta As String
        
    ruta = getRutaRaiz()
    MkDir ruta
    
    ruta = getRutaTemporal()
    MkDir ruta
    'SetAttr myPath, vbHidden 'Ocultamos la  carpeta "Temp"
    
    Call vaciarTemporal
End Sub

' Vacía la carpeta temporal, al guardar o deshacer
'
Private Sub vaciarTemporal()
    On Error Resume Next
    Kill agregarRuta(getRutaTemporal(), "*.*")
End Sub

Private Function getRutaRaiz() As String
    If strCarpetaRaiz = "" Then
        getRutaRaiz = agregarRuta(Application.CurrentProject.Path, CARPETA_RAIZ)
    Else
        getRutaRaiz = strCarpetaRaiz
    End If
End Function
Private Function getRutaTemporal() As String
    getRutaTemporal = agregarRuta(getRutaRaiz(), CARPETA_TEMPORAL)
End Function
Private Function getRutaTabla() As String
    getRutaTabla = agregarRuta(getRutaRaiz(), strCarpetaDatos)
End Function
Private Function getRutaId(id) As String
    getRutaId = agregarRuta(getRutaTabla(), id)
End Function

Private Function buscarRutaArchivo(ByVal id As Variant, ByVal myFile As Variant) As String
On Error GoTo Errores
    Dim myPaths(1 To 2) As String
    Dim myPath As String
    Dim i As Integer
    myPaths(1) = getRutaTemporal()
    myPaths(2) = getRutaId(id)
    'Buscamos en cual está el archivo
    For i = LBound(myPaths) To UBound(myPaths)
        myPath = myPaths(i)
        If Dir(myPath, vbDirectory) <> "" Then
            If Dir(agregarRuta(myPath, myFile)) <> "" Then
                buscarRutaArchivo = myPath
                Exit Function
            End If
        End If
Siguiente:
    Next
    buscarRutaArchivo = ""
Salida:
    Exit Function
Errores:
    Select Case Err.Number
        Case 53 'Archivo no encontrado
            Resume Siguiente
        Case 490
            MsgBox "El archivo """ & myFile & """ no existe.", vbInformation + vbOKOnly, "ERROR"
        Case Else
            Call mostrarError("buscarRutaArchivo")
    End Select
    Resume Salida
End Function

Private Function obtenerCarpetaTabla() As String
    Dim td As DAO.TableDef
    Dim ruta As String
    For Each td In CurrentDb.TableDefs
        If td.Name = strNombreTabla Then
            If td.Connect <> "" Then
                ruta = Replace(td.Connect, ";DATABASE=", "")
                ruta = extraerRuta(ruta)
                ruta = agregarRuta(ruta, CARPETA_RAIZ)
                obtenerCarpetaTabla = ruta
                Exit Function
            Else
                Exit For
            End If
        End If
    Next
    ruta = agregarRuta(Application.CurrentProject.Path, CARPETA_RAIZ)
    obtenerCarpetaTabla = ruta
End Function

Public Function agregarRuta(ByVal myPath As String, ParamArray items()) As String
    Dim item
    For Each item In items
        If Right(myPath, 1) <> "\" And Left(item, 1) <> "\" Then
            myPath = myPath & "\" & item
        Else
            myPath = myPath & item
        End If
    Next
    agregarRuta = myPath
End Function

Private Function extraerRuta(ByVal ruta As String) As String
    Dim posicion As Integer
    posicion = InStrRev(ruta, "\")
    If posicion > 1 Then ruta = Left(ruta, posicion - 1)
    extraerRuta = ruta & "\"
End Function

' Abrir ventana de diálogo y buscar archivos para adjuntar a documentación
'
Private Function dialogoSeleccionarArchivo() As String
    On Error GoTo Errores
    
    Dim fDialog As Office.FileDialog
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .ButtonName = "Seleccionar"
        .Title = "Seleccionar el archivo"
        .InitialFileName = Application.CurrentProject.Path & "\"
        .InitialView = msoFileDialogViewDetails
        .Filters.Clear
        .Filters.Add "Todos los Archivos", "*.*"
        If .Show = True Then
            dialogoSeleccionarArchivo = .SelectedItems(1)
        Else
            'No hacemos nada
        End If
    End With
Salida:
    Set fDialog = Nothing
    Exit Function
Errores:
    Call mostrarError("dialogoSeleccionarArchivo")
    Resume Salida
End Function


' Abrir ventana de diálogo y elegir la carpeta donde exportar la documentación
'
Private Function dialogoSeleccionarCarpeta() As String
    On Error GoTo Errores
    
    Dim fDialog As Office.FileDialog
    
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Selecciona la carpeta donde se va a exportar el archivo"
        .ButtonName = "Aceptar"
        .InitialFileName = Application.CurrentProject.Path & "\"
        .InitialView = msoFileDialogViewList
        If .Show = -1 Then
            dialogoSeleccionarCarpeta = CStr(.SelectedItems.item(1))
        Else
            'No hacemos nada
        End If
    End With
Salida:
    Set fDialog = Nothing
    Exit Function
Errores:
    Call mostrarError("dialogoSeleccionarCarpeta")
    Resume Salida
End Function


Private Sub mostrarError(ByVal nombreFunc As String)
    MsgBox "Se ha producido el error: " & Err.Number & " - " & Err.Description & " en la función " & nombreFunc, _
        vbInformation + vbOKOnly, "ERROR"
End Sub

