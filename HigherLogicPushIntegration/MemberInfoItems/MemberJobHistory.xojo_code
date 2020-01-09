#tag Class
Protected Class MemberJobHistory
Inherits JSONClass
	#tag Property, Flags = &h0
		City As string
	#tag EndProperty

	#tag Property, Flags = &h0
		CompanyName As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Country As string
	#tag EndProperty

	#tag Property, Flags = &h0
		EndDate As Date
	#tag EndProperty

	#tag Property, Flags = &h0
		EndDateString As string
	#tag EndProperty

	#tag Property, Flags = &h0
		StartDate As Date
	#tag EndProperty

	#tag Property, Flags = &h0
		StartDateString As string
	#tag EndProperty

	#tag Property, Flags = &h0
		State As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Title As string
	#tag EndProperty

	#tag Property, Flags = &h0
		WebsiteUrl As string
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="City"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CompanyName"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Country"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EndDateString"
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
			Name="StartDateString"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="State"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Title"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="WebsiteUrl"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
