#tag Class
Protected Class CommunityGroup
Inherits JSONClass
	#tag Property, Flags = &h0
		BeginDate As String
	#tag EndProperty

	#tag Property, Flags = &h0
		EndDate As String
	#tag EndProperty

	#tag Property, Flags = &h0
		GroupKey As string
	#tag EndProperty

	#tag Property, Flags = &h0
		GroupName As string
	#tag EndProperty

	#tag Property, Flags = &h0
		GroupSubType As string
	#tag EndProperty

	#tag Property, Flags = &h0
		GroupType As string
	#tag EndProperty

	#tag Property, Flags = &h0
		ParentCommunityKey As string
	#tag EndProperty

	#tag Property, Flags = &h0
		RoleDescription As string
	#tag EndProperty

	#tag Property, Flags = &h0
		SinceDate As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="GroupKey"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="GroupName"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="GroupSubType"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="GroupType"
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
			Name="ParentCommunityKey"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="RoleDescription"
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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="BeginDate"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="EndDate"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="SinceDate"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
