#tag WebPage
Begin WebPage frmUpdateInfo
   Compatibility   =   ""
   Cursor          =   0
   Enabled         =   True
   Height          =   482
   HelpTag         =   ""
   HorizontalCenter=   0
   ImplicitInstance=   True
   Index           =   0
   IsImplicitInstance=   False
   Left            =   0
   LockBottom      =   False
   LockHorizontal  =   False
   LockLeft        =   False
   LockRight       =   False
   LockTop         =   False
   LockVertical    =   False
   MinHeight       =   482
   MinWidth        =   950
   Style           =   "149482928"
   TabOrder        =   0
   Title           =   "Update Information"
   Top             =   0
   VerticalCenter  =   0
   Visible         =   True
   Width           =   940
   ZIndex          =   1
   _DeclareLineRendered=   False
   _HorizontalPercent=   0.0
   _ImplicitInstance=   False
   _IsEmbedded     =   False
   _Locked         =   False
   _NeedsRendering =   True
   _OfficialControl=   False
   _OpenEventFired =   False
   _ShownEventFired=   False
   _VerticalPercent=   0.0
   Begin conMemInfo conMemInfo1
      Cursor          =   0
      Enabled         =   True
      Height          =   415
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   16
      lnErrorCount    =   -1
      LockBottom      =   True
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      LockVertical    =   False
      mbIsCPD         =   False
      msEmail         =   ""
      msNameSuffix    =   ""
      msRegion        =   ""
      Scope           =   0
      ScrollbarsVisible=   0
      Style           =   "0"
      TabOrder        =   0
      Top             =   47
      VerticalCenter  =   0
      Visible         =   True
      Width           =   910
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _ShownEventFired=   False
      _VerticalPercent=   0.0
   End
   Begin WebLabel Label1
      Cursor          =   1
      Enabled         =   True
      HasFocusRing    =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Multiline       =   False
      Scope           =   0
      Style           =   "866787714"
      TabOrder        =   1
      Text            =   "Your account data as ASPE has it:"
      Top             =   14
      VerticalCenter  =   0
      Visible         =   True
      Width           =   910
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin WebLabel Label2
      Cursor          =   1
      Enabled         =   True
      HasFocusRing    =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   364
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Multiline       =   False
      Scope           =   0
      Style           =   "1304282530"
      TabOrder        =   1
      Text            =   "To edit your account press the edit button:"
      Top             =   418
      VerticalCenter  =   0
      Visible         =   True
      Width           =   303
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
   Begin SMTPServerASPE SMTPServer
      CertificateFile =   
      CertificatePassword=   ""
      CertificateRejectionFile=   
      ConnectionType  =   2
      Height          =   32
      Index           =   -2147483648
      Left            =   80
      LockedInPosition=   False
      Scope           =   0
      Secure          =   False
      SMTPConnectionMode=   0
      Style           =   "-1"
      TabPanelIndex   =   0
      Top             =   80
      Width           =   32
   End
   Begin WebLabel lblVersion
      Cursor          =   1
      Enabled         =   True
      HasFocusRing    =   True
      Height          =   22
      HelpTag         =   ""
      HorizontalCenter=   0
      Index           =   -2147483648
      Left            =   16
      LockBottom      =   False
      LockedInPosition=   False
      LockHorizontal  =   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      LockVertical    =   False
      Multiline       =   False
      Scope           =   0
      Style           =   "-1"
      TabOrder        =   2
      Text            =   ""
      Top             =   460
      VerticalCenter  =   0
      Visible         =   True
      Width           =   100
      ZIndex          =   1
      _DeclareLineRendered=   False
      _HorizontalPercent=   0.0
      _IsEmbedded     =   False
      _Locked         =   False
      _NeedsRendering =   True
      _OfficialControl=   False
      _OpenEventFired =   False
      _VerticalPercent=   0.0
   End
End
#tag EndWebPage

#tag WindowCode
	#tag Event
		Sub Shown()
		  lblVersion.Text = "Build:" + Str(App.NonReleaseVersion)
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Function CreateMsg() As String
		  
		  Dim lsMsg as String
		  
		  
		  
		  lsMsg = EndOfLine + EndOfLine
		  
		  'lsMsg = lsMsg + "This is a test: This message will be sent automaticaly to the chapter officers of the Previous Chapter " + EndOfLine
		  'lsMsg = lsMsg + msPrevPresEmail + " & " + msPrevVPMEmail + EndOfLine + EndOfLine
		  'lsMsg = lsMsg + "Also to the chapter officers of the New Chapter: " + EndOfLine
		  'lsMsg = lsMsg + msNewPresEmail + " & " + msNewVPMEmail + EndOfLine + EndOfLine
		  'lsMsg = lsMsg + "Start of actual message:"  + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "This is an automated email from ASPE."  + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "To Chapter President and Vice President Membership"  + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "This email is to inform you that " + conMemInfo1.txtFirst.Text + " " + conMemInfo1.txtLast.Text + " has changed  his chapter from " + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "Previous Chapter: " + msPreviousChapter + EndOfLine
		  lsMsg = lsMsg + "New Chapter: " + conMemInfo1.cboChapterName.Text  + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "It is suggested that the officers from " + conMemInfo1.cboChapterName.Text  + " welcome him. "
		  
		  lsMsg = lsMsg + "Their email is: " + conMemInfo1.txtPrimaryEmail.Text + " and their phone is: " + conMemInfo1.txtPPhone.Text  + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "Thank You for your support!" + EndOfLine + EndOfLine
		  
		  lsMsg = lsMsg + "ASPE Membership" + EndOfLine
		  lsMsg = lsMsg + "847-296-0002"  + EndOfLine + EndOfLine
		  
		  Return lsMsg
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendNotification()
		  Dim Msg as New EmailMessage
		  
		  
		  SMTPServer = New SMTPServer
		  
		  SMTPServer.Address = "localhost" 'csBulkMailSMTPServer
		  SMTPServer.Port = 25 'cnBulkEmailPort
		  SMTPServer.Username = ""   'csBulkMailSMTPUserID
		  SMTPServer.Password = ""   'csBulkEmailSMTPPassword
		  Msg.FromAddress = "Membership@aspe.org"
		  Msg.AddCCRecipient "Membership@aspe.org"
		  
		  if msPrevPresEmail <> "" Then
		    Msg.AddRecipient msPrevPresEmail
		  end
		  if msPrevVPMEmail <> "" Then
		    Msg.AddRecipient msPrevVPMEmail
		  end
		  if msNewPresEmail <> "" Then
		    Msg.AddRecipient msNewPresEmail
		  end
		  if msNewVPMEmail <> "" Then
		    Msg.AddRecipient msNewVPMEmail
		  end
		  'Msg.AddRecipient "Rich@RAlbrecht.net"
		  
		  'System.DebugLog("Sending email to: " + csEmailAddress)
		  
		  Msg.subject = "Notice of Chapter Affiliation change for " + conMemInfo1.txtFirst.Text + ", " + conMemInfo1.txtLast.Text
		  Msg.BodyPlainText = CreateMsg()
		  SMTPServer.Messages.Append( Msg)
		  SMTPServer.SendMail
		  
		  System.DebugLog("Last Error: " + Str(SMTPServer.LastErrorCode) )
		  
		  
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		mbIsMember As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		msCurrentChapter As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msNewChapterCode As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msNewPresEmail As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msNewVPMEmail As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msPreviousChapter As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msPreviousChapterCode As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msPrevPresEmail As String
	#tag EndProperty

	#tag Property, Flags = &h0
		msPrevVPMEmail As String
	#tag EndProperty


#tag EndWindowCode

#tag ViewBehavior
	#tag ViewProperty
		Name="Cursor"
		Visible=true
		Group="Behavior"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Auto"
			"1 - Standard Pointer"
			"2 - Finger Pointer"
			"3 - IBeam"
			"4 - Wait"
			"5 - Help"
			"6 - Arrow All Directions"
			"7 - Arrow North"
			"8 - Arrow South"
			"9 - Arrow East"
			"10 - Arrow West"
			"11 - Arrow North East"
			"12 - Arrow North West"
			"13 - Arrow South East"
			"14 - Arrow South West"
			"15 - Splitter East West"
			"16 - Splitter North South"
			"17 - Progress"
			"18 - No Drop"
			"19 - Not Allowed"
			"20 - Vertical IBeam"
			"21 - Crosshair"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Enabled"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Behavior"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HelpTag"
		Visible=true
		Group="Behavior"
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HorizontalCenter"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Index"
		Group="ID"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="IsImplicitInstance"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Left"
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockBottom"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockHorizontal"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockLeft"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockRight"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockTop"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LockVertical"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Behavior"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Behavior"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="TabOrder"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Behavior"
		InitialValue="Untitled"
		Type="String"
		EditorType="MultiLineEditor"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Top"
		Group="Position"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="VerticalCenter"
		Group="Behavior"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Behavior"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ZIndex"
		Group="Behavior"
		InitialValue="1"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_HorizontalPercent"
		Group="Behavior"
		Type="Double"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_ImplicitInstance"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_IsEmbedded"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_Locked"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_NeedsRendering"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_OfficialControl"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_ShownEventFired"
		Group="Behavior"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="_VerticalPercent"
		Group="Behavior"
		Type="Double"
	#tag EndViewProperty
#tag EndViewBehavior
