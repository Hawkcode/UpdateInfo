#tag Class
Protected Class MemberInfo
Inherits JSONClass
	#tag Property, Flags = &h0
		CommunityGroups() As CommunityGroup
	#tag EndProperty

	#tag Property, Flags = &h0
		Demographics() As MemberDemographic
	#tag EndProperty

	#tag Property, Flags = &h0
		Education() As MemberEducation
	#tag EndProperty

	#tag Property, Flags = &h0
		JobHistory() As MemberJobHistory
	#tag EndProperty

	#tag Property, Flags = &h0
		MemberDetails As MemberDetails
	#tag EndProperty

	#tag Property, Flags = &h0
		SecurityGroups() As SecurityGroup
	#tag EndProperty


	#tag ViewBehavior
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
