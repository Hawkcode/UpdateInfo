#tag Class
Protected Class PushMemberInfoResponse
Inherits JSONClass
	#tag Method, Flags = &h0
		Sub constructor(JSONString As String)
		  Dim res As New JSONItem(JSONString)
		  
		  If res.HasName("count") Then count=res.Value("count")
		  If res.HasName("errorMessage") Then errorMessage=res.Value("errorMessage")
		  if res.HasName("success") then success=res.Value("success")
		End Sub
	#tag EndMethod


	#tag Note, Name = About
		
		This class is returned by PushMemberInfoRequest.Send_Request and contains the errorMessage, records count and success flag.
		
	#tag EndNote


	#tag Property, Flags = &h0
		count As integer
	#tag EndProperty

	#tag Property, Flags = &h0
		errorMessage As string
	#tag EndProperty

	#tag Property, Flags = &h0
		success As boolean
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="count"
			Group="Behavior"
			Type="integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="errorMessage"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="success"
			Group="Behavior"
			Type="boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
