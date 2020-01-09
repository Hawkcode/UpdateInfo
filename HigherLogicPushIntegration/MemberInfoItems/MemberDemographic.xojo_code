#tag Class
Protected Class MemberDemographic
Inherits JSONClass
	#tag Property, Flags = &h0
		DemographicKey As string
	#tag EndProperty

	#tag Property, Flags = &h0
		DemographicTypeKey As string
	#tag EndProperty

	#tag Property, Flags = &h0
		DemographicTypeValue As string
	#tag EndProperty

	#tag Property, Flags = &h0
		DemographicValue As string
	#tag EndProperty

	#tag Property, Flags = &h0
		IsFreeForm As boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		isFreeFormString As string
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="DemographicKey"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DemographicTypeKey"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DemographicTypeValue"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DemographicValue"
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
			Name="IsFreeForm"
			Group="Behavior"
			Type="boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="isFreeFormString"
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
