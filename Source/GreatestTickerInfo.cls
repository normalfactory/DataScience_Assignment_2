VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GreatestTickerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Purpose : Stores the greatest ticker information while processing the worksheet


' Name of the ticker that had the greatest increase
Private m_greatestIncreaseTicker As String

' Value of the greatest increase
Private m_greatestIncreaseValue As Double

' Name of the ticker that had the greatest decrease
Private m_greatestDecreaseTicker As String

' Value of the greatest decrease
Private m_greatestDecreaseValue As Double

' Name of the ticker that had the greatest volumne increase
Private m_greatestVolumeTicker As String

' Value of the greatest volume for the year
Private m_greatestVolumeValue As Double



'- Constructor
Private Sub Class_Initialize()
' Purpose : Creates a new instance and sets the default values

    m_greatestIncreaseTicker = ""
    m_greatestIncreaseValue = 0

    m_greatestDecreaseTicker = ""
    m_greatestDecreaseValue = 0
    
    m_greatestVolumeTicker = ""
    m_greatestVolumeValue = 0
End Sub



'- Properties
Public Property Get GreatestDecreaseTicker() As String
    GreatestDecreaseTicker = m_greatestDecreaseTicker
End Property

Public Property Let GreatestDecreaseTicker(value As String)
    m_greatestDecreaseTicker = value
End Property


Public Property Get GreatestDecreaseValue() As Double
    GreatestDecreaseValue = m_greatestDecreaseValue
End Property

Public Property Let GreatestDecreaseValue(value As Double)
    m_greatestDecreaseValue = value
End Property


Public Property Get GreatestIncreaseTicker() As String
    GreatestIncreaseTicker = m_greatestIncreaseTicker
End Property

Public Property Let GreatestIncreaseTicker(value As String)
    m_greatestIncreaseTicker = value
End Property

Public Property Get GreatestIncreaseValue() As Double
    GreatestIncreaseValue = m_greatestIncreaseValue
End Property

Public Property Let GreatestIncreaseValue(value As Double)
    m_greatestIncreaseValue = value
End Property


Public Property Get GreatestVolumeTicker() As String
    GreatestVolumeTicker = m_greatestVolumeTicker
End Property

Public Property Let GreatestVolumeTicker(value As String)
    m_greatestVolumeTicker = value
End Property

Public Property Get GreatestVolumeValue() As Double
    GreatestVolumeValue = m_greatestVolumeValue
End Property

Public Property Let GreatestVolumeValue(value As Double)
    m_greatestVolumeValue = value
End Property


