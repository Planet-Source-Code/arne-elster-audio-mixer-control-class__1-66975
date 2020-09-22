VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Windows Mixer Class"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cboDev 
      Height          =   315
      Left            =   1350
      Style           =   2  'Dropdown-Liste
      TabIndex        =   10
      Top             =   75
      Width           =   4065
   End
   Begin VB.ComboBox cboDevMode 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   75
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown-Liste
      TabIndex        =   9
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "WaveOut Balance"
      Height          =   390
      Left            =   2850
      TabIndex        =   8
      Top             =   4725
      Width           =   2040
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Speakers Balance"
      Height          =   390
      Left            =   600
      TabIndex        =   7
      Top             =   4725
      Width           =   2040
   End
   Begin VB.CommandButton Command6 
      Caption         =   "WaveOut to 100% Volume"
      Height          =   390
      Left            =   600
      TabIndex        =   6
      Top             =   4275
      Width           =   2040
   End
   Begin VB.CommandButton Command5 
      Caption         =   "mute speakers"
      Height          =   390
      Left            =   2850
      TabIndex        =   5
      Top             =   4275
      Width           =   2040
   End
   Begin VB.CommandButton Command4 
      Caption         =   "mute CD Player"
      Height          =   390
      Left            =   2857
      TabIndex        =   4
      Top             =   3825
      Width           =   2040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select WaveIn Mic"
      Height          =   390
      Left            =   607
      TabIndex        =   3
      Top             =   3825
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Everything to 100% volume"
      Height          =   390
      Left            =   2857
      TabIndex        =   2
      Top             =   3375
      Width           =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Everything to 50% volume"
      Height          =   390
      Left            =   607
      TabIndex        =   1
      Top             =   3375
      Width           =   2040
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   5340
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsMix  As clsWinMixer
Private lngMode As Long

Private Sub cboDev_Click()
    clsMix.DeviceClose

    If Not clsMix.DeviceOpen(cboDev.ListIndex, lngMode) Then
        MsgBox "Couldn't open the device!", vbExclamation
        Exit Sub
    End If

    ShowMixerStuff
End Sub

Private Sub cboDevMode_Click()
    Dim i       As Long

    Select Case cboDevMode.ListIndex
        Case 0: lngMode = MIXER_OPENBY_WAVEIN_ID
        Case 1: lngMode = MIXER_OPENBY_WAVEOUT_ID
        Case 2: lngMode = MIXER_OPENBY_MIDIIN_ID
        Case 3: lngMode = MIXER_OPENBY_MIDIOUT_ID
        Case 4: lngMode = MIXER_OPENBY_MIXER_ID
    End Select

    clsMix.DeviceClose
    cboDev.Clear
    List1.Clear

    For i = 0 To clsMix.DeviceCount(lngMode) - 1
        cboDev.AddItem clsMix.DeviceName(i, lngMode)
    Next

    If cboDev.ListCount > 0 Then
        cboDev.ListIndex = 0
    End If
End Sub

Private Sub Command1_Click()
    Dim i   As Long
    Dim j   As Long

    ' we want to change the volume of every destination
    ' and every source connected with a destination

    ' go through every destination
    For i = 0 To clsMix.DestinationCount - 1
        ' set the volume of the destination to 50%
        ' for every channel (-1)
        clsMix.DestinationVolume(i, -1) = 50

        ' go through every source connected to the destination
        For j = 0 To clsMix.SourceCount(i) - 1
            ' set the volume of the source to 50% for every channel (-1)
            clsMix.SourceVolume(i, j, -1) = 50
        Next
    Next

    ShowMixerStuff
End Sub

Private Sub Command2_Click()
    Dim i   As Long
    Dim j   As Long

    ' we want to change the volume of every destination
    ' and every source connected with a destination

    ' go through every destination
    For i = 0 To clsMix.DestinationCount - 1
        ' set the volume of the destination to 100%
        ' for every channel (-1)
        clsMix.DestinationVolume(i, -1) = 100

        ' go through every source connected to the destination
        For j = 0 To clsMix.SourceCount(i) - 1
            ' set the volume of the source to 100% for every channel (-1)
            clsMix.SourceVolume(i, j, -1) = 100
        Next
    Next

    ShowMixerStuff
End Sub

Private Sub Command3_Click()
    Dim i   As Long
    Dim j   As Long

    ' we want to select the microphone as the
    ' recording source for the WaveIn
    
    ' go through every destination
    For i = 0 To clsMix.DestinationCount - 1
        ' if the destination's type is WaveIn (recording)...
        If clsMix.DestinationType(i) = MIXERLINE_DST_WAVEIN Then

            ' ... search for a connected microphone source
            For j = 0 To clsMix.SourceCount(i) - 1
                If clsMix.SourceType(i, j) = MIXERLINE_SRC_MICROPHONE Then
                    ' found it, select it!
                    clsMix.SourceSelected(i, j) = True
                    Exit For
                End If
            Next

        End If
    Next

    ShowMixerStuff
End Sub

Private Sub Command4_Click()
    Dim i   As Long
    Dim j   As Long

    ' we want to mute or demute the CD source connected to the speakers

    ' go through every destination
    For i = 0 To clsMix.DestinationCount - 1
        ' if the destination's type is Speakers...
        If clsMix.DestinationType(i) = MIXERLINE_DST_SPEAKERS Then

            ' ... search for a CD source connected to the destination
            For j = 0 To clsMix.SourceCount(i) - 1
                If clsMix.SourceType(i, j) = MIXERLINE_SRC_COMPACTDISC Then
                    ' mute or demute
                    clsMix.SourceMute(i, j) = Not clsMix.SourceMute(i, j)
                    Exit For
                End If
            Next

        End If
    Next

    ShowMixerStuff
End Sub

Private Sub Command5_Click()
    Dim i   As Long

    ' we want to mute the speakers destination

    ' go through the destinations
    For i = 0 To clsMix.DestinationCount - 1
        ' if the destination's type is Speakers
        If clsMix.DestinationType(i) = MIXERLINE_DST_SPEAKERS Then
            ' mute or demute it
            clsMix.DestinationMute(i) = Not clsMix.DestinationMute(i)
        End If
    Next

    ShowMixerStuff
End Sub

Private Sub Command6_Click()
    Dim i   As Long
    Dim j   As Long

    ' we want to change the volume of the WaveOut source
    ' connected to the Speakers to 100%

    ' go through all destinations
    For i = 0 To clsMix.DestinationCount - 1
        ' if the destination's type is Speakers
        If clsMix.DestinationType(i) = MIXERLINE_DST_SPEAKERS Then

            ' go through every source connected to the destination
            For j = 0 To clsMix.SourceCount(i) - 1
                ' if the source's type is WaveOut
                If clsMix.SourceType(i, j) = MIXERLINE_SRC_WAVEOUT Then
                    ' change the volume for all channels (-1) to 100%
                    clsMix.SourceVolume(i, j, -1) = 100
                    Exit For
                End If
            Next

        End If
    Next

    ShowMixerStuff
End Sub

Private Sub Command7_Click()
    Dim i   As Long

    ' we want to change the balance of the Speakers destination

    ' go through every destination
    For i = 0 To clsMix.DestinationCount - 1
        ' if destination's type is Speakers
        If clsMix.DestinationType(i) = MIXERLINE_DST_SPEAKERS Then
            ' if the volume of the left channel (0) was 25%
            If clsMix.DestinationVolume(i, 0) = 25 Then
                ' set the volume of the left channel (0) to 50%
                clsMix.DestinationVolume(i, 0) = 50

                ' if the destination got more then 1 channel
                If clsMix.DestinationChannels(i) > 1 Then
                    ' set the right channel's (1) volume to 50%
                    ' to set the balance to 0
                    clsMix.DestinationVolume(i, 1) = 50
                End If
            Else
                ' set the volume of the left channel (0) to 25%
                clsMix.DestinationVolume(i, 0) = 25

                ' if the destination got more then 1 channel
                If clsMix.DestinationChannels(i) > 1 Then
                    ' set the right channel's (1) volume to
                    ' 75% to change the balance
                    clsMix.DestinationVolume(i, 1) = 75
                End If
            End If
        End If
    Next

    ShowMixerStuff
End Sub

Private Sub Command8_Click()
    Dim i   As Long
    Dim j   As Long

    ' we want to change the balance of the WaveOut source
    ' connected to the Speakers destination

    ' check all destinations
    For i = 0 To clsMix.DestinationCount - 1
        ' if the destination's type is Speakers ...
        If clsMix.DestinationType(i) = MIXERLINE_DST_SPEAKERS Then

            ' ... check all sources connected to the destinations
            For j = 0 To clsMix.SourceCount(i) - 1

                ' if the source's type is WaveOut change its balance
                If clsMix.SourceType(i, j) = MIXERLINE_SRC_WAVEOUT Then
                    ' if the volume of the left channel (0) is 25%
                    If clsMix.SourceVolume(i, j, 0) = 25 Then
                        ' set the volume to 50%
                        clsMix.SourceVolume(i, j, 0) = 50

                        ' if source got more then 1 channel...
                        If clsMix.SourceChannels(i, j) > 1 Then
                            ' set the right channel's (1) volume to 50%
                            ' so the balance is 0
                            clsMix.SourceVolume(i, j, 1) = 50
                        End If
                    Else
                        ' set the volume of the left channel to 25%
                        clsMix.SourceVolume(i, j, 0) = 25

                        ' if the source got more then 1 channel...
                        If clsMix.SourceChannels(i, j) > 1 Then
                            ' set the right channel's volume (1) to 75%
                            clsMix.SourceVolume(i, j, 1) = 75
                        End If
                    End If
                End If

            Next

        End If
    Next

    ShowMixerStuff
End Sub

Private Sub Form_Load()
    Set clsMix = New clsWinMixer

    cboDevMode.ListIndex = 4
End Sub

Private Sub ShowMixerStuff()
    Dim i   As Long
    Dim j   As Long

    List1.Clear

    For i = 0 To clsMix.DestinationCount() - 1
        List1.AddItem "Destination: " & clsMix.DestinationName(i) & " (Typ: " & clsMix.DestinationType(i) & ")"
        List1.AddItem "  Volume: " & clsMix.DestinationVolume(i)
        List1.AddItem "  Muted: " & clsMix.DestinationMute(i)
        List1.AddItem "  Channels: " & clsMix.DestinationChannels(i)
        List1.AddItem ""

        For j = 0 To clsMix.SourceCount(i) - 1
            List1.AddItem "   Source: " & clsMix.SourceName(i, j) & " (Typ: " & clsMix.SourceType(i, j) & ")"
            List1.AddItem "       Volume: " & clsMix.SourceVolume(i, j)
            List1.AddItem "       Muted: " & clsMix.SourceMute(i, j)
            List1.AddItem "       Selected: " & clsMix.SourceSelected(i, j)
            List1.AddItem "       Channels: " & clsMix.SourceChannels(i, j)
        Next

        List1.AddItem ""
    Next
End Sub
