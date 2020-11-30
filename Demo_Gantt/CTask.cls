VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDuracion As Double
Private pID As Long
Private pDescripcion As String
Private pLeft As Double
Private pTop As Double
Private pWidth As Double
Private pHeight As Double
Private pColor As Long
Private pIDPrecedente As Variant
Private pDuracionPrevias As Double

'Duracion property
Public Property Get duracion() As Double
    duracion = pDuracion
End Property
Public Property Let duracion(value As Double)
    pDuracion = value
End Property

'ID property
Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(value As Long)
    pID = value
End Property

'Descripcion property
Public Property Get Descripcion() As String
    Descripcion = pDescripcion
End Property
Public Property Let Descripcion(value As String)
    pDescripcion = value
End Property

'Left property
Public Property Get Left() As Double
    Left = pLeft
End Property

Public Property Let Left(value As Double)
    pLeft = value
End Property

'Top property
Public Property Get Top() As Double
    Top = pTop
End Property
Public Property Let Top(value As Double)
    pTop = value
End Property

'Height property
Public Property Get Height() As Double
    Height = pHeight
End Property
Public Property Let Height(value As Double)
    pHeight = value
End Property

'Color property
Public Property Get Color() As Double
    Color = pColor
End Property
Public Property Let Color(value As Double)
    pColor = value
End Property

'ID_Precedente property
Public Property Get ID_Precedente() As Long
    ID_Precedente = pIDPrecedente
End Property
Public Property Let ID_Precedente(value As Long)
    pIDPrecedente = value
End Property

'pDuracionPrevias property
Public Property Get DuracionPrevias() As Double
    DuracionPrevias = pDuracionPrevias
End Property
Public Property Let DuracionPrevias(value As Double)
    pDuracionPrevias = value
End Property

'Width property
Public Property Let Width(value As Double)
    pWidth = value
End Property

Public Property Get Width() As Double
    Width = pWidth
End Property


