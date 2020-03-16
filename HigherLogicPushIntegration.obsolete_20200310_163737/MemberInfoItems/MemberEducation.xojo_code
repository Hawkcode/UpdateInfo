#tag Class
Protected Class MemberEducation
Inherits JSONClass
	#tag Property, Flags = &h0
		Advisor As string
	#tag EndProperty

	#tag Property, Flags = &h0
		City As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Country As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Degree As string
	#tag EndProperty

	#tag Property, Flags = &h0
		DegreeYear As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		DegreeYearString As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Dissertation As string
	#tag EndProperty

	#tag Property, Flags = &h0
		FieldOfStudy As string
	#tag EndProperty

	#tag Property, Flags = &h0
		FromYear As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		FromYearString As string
	#tag EndProperty

	#tag Property, Flags = &h0
		IsHighestDegreeAttained As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		IsHighestDegreeAttainedString As string
	#tag EndProperty

	#tag Property, Flags = &h0
		SchoolName As string
	#tag EndProperty

	#tag Property, Flags = &h0
		State As string
	#tag EndProperty

	#tag Property, Flags = &h0
		ToYear As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		ToYearString As string
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Advisor"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="City"
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
			Name="Degree"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DegreeYearString"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Dissertation"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FieldOfStudy"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FromYearString"
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
			Name="IsHighestDegreeAttainedString"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
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
			Name="SchoolName"
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
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ToYearString"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DegreeYear"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="FromYear"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="IsHighestDegreeAttained"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ToYear"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
