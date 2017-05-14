VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FormularioVinculado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private misAdjuntos As ClaseArchivosAdjuntos

Private Sub Form_Load()
    Set misAdjuntos = New ClaseArchivosAdjuntos
    misAdjuntos.NombreTabla = "TablaVinculada"
    misAdjuntos.CampoId = Me.id
    misAdjuntos.CampoAdjuntos = Me.adjuntos
    misAdjuntos.ComboboxIndependiente = Me.comboDocumento
    misAdjuntos.Inicializar
End Sub

Private Sub Form_AfterInsert()
    Call misAdjuntos.Procesar
End Sub

Private Sub Form_AfterUpdate()
    Call misAdjuntos.Procesar
End Sub

Private Sub Form_Current()
    Call misAdjuntos.Actualizar
    Call misAdjuntos.Asear
End Sub

Private Sub cmdDocument_Click()
    Call misAdjuntos.BotonElegir
End Sub

Private Sub cmdVerDoc_Click()
    Call misAdjuntos.BotonVer
End Sub

Private Sub cmdExportaDoc_Click()
    Call misAdjuntos.BotonExportar
End Sub

Private Sub cmdBorrarDoc_Click()
    Call misAdjuntos.BotonBorrar
End Sub
