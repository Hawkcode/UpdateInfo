#tag Module
Protected Module mSmartSQL
	#tag Enum, Name = eClauseOp, Type = Integer, Flags = &h0
		Equal
		  Like
		  GreaterThan
		  GreaterThanOrEqual
		  LessThan
		  LessThanOrEqual
		  NotEqual
		  StartWith
		EndWith
	#tag EndEnum

	#tag Enum, Name = eJoinType, Type = Integer, Flags = &h0
		Inner
		  Left
		  Right
		Full
	#tag EndEnum

	#tag Enum, Name = eStatementType, Type = Integer, Flags = &h0
		Type_Select
		  Type_Insert
		  Type_Update
		Type_Delete
	#tag EndEnum

	#tag Enum, Name = eWhereClauseLogic, Type = Integer, Flags = &h0
		Logic_And
		Logic_Or
	#tag EndEnum


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
End Module
#tag EndModule
