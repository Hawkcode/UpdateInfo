#tag Class
Class cSmartSQL
	#tag Method, Flags = &h0
		Sub AddComplexWhereClause(sClause as string)
		  AddComplexWhereClause(sClause, eWhereClauseLogic.Logic_And)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddComplexWhereClause(sClause as string, logic as eWhereClauseLogic)
		  if m_arsWhereClause.Ubound+1 > 0 then
		    m_arsWhereClause.Append( GetLogicOp(logic) )
		  end if
		  m_arsWhereClause.Append("(")
		  m_arsWhereClause.Append(sClause)
		  m_arsWhereClause.Append(")")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddField(sFieldName as string, sTableName as string = "")
		  m_arsFieldName.Append( FieldName(sFieldName,sTableName, AutoBracket) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddFields(ParamArray arsField() as string)
		  'Pass the argument array to the array handler
		  AddFields(arsField,"")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddFields(arsField() as string, sTableName as string = "")
		  for each sField as string in arsField
		    AddField(sField,sTableName)
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddOrderClause(sFieldName as string, bDesc as boolean = false, sTableName as string = "")
		  if bDesc then
		    m_arsOrderBy.Append( FieldName(sFieldName,sTableName,AutoBracket) + " DESC" )
		  else
		    m_arsOrderBy.Append( FieldName(sFieldName,sTableName,AutoBracket) )
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddSimpleWhereClause(sFieldName as string, Value as Variant)
		  AddSimpleWhereClause(sFieldName, Value, "", eClauseOp.Equal, eWhereClauseLogic.Logic_And)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddSimpleWhereClause(sFieldName as string, Value as Variant, sTableName as string)
		  AddSimpleWhereClause(sFieldName,Value,sTableName, eClauseOp.Equal, eWhereClauseLogic.Logic_And)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddSimpleWhereClause(sFieldName as string, Value as Variant, sTableName as string, clause as eClauseOp)
		  AddSimpleWhereClause(sFieldName,Value,sTableName,clause, eWhereClauseLogic.Logic_And)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddSimpleWhereClause(sFieldName as string, Value as Variant, sTableName as string, clause as eClauseOp, logic as eWhereClauseLogic)
		  if clause = eClauseOp.Like and VarType(Value)<>Variant.TypeString then
		    clause = eClauseOp.Equal
		  elseif clause = eClauseOp.StartWith or clause = eClauseOp.EndWith then
		    'if the value isn't a string we'll coerce it into one, since there
		    'aren't equivalents to StartWith and EndWith for non-string values.
		    Value = CStr(Value)
		  end if
		  
		  if Value = nil then
		    raise new cInvalidValueException
		  end if
		  
		  if AutoLike then
		    if clause = eClauseOp.Like then
		      Value = kLikeWildcard + CStr(Value) + kLikeWildcard
		    elseif clause = eClauseOp.StartWith then
		      Value = CStr(Value) + kLikeWildcard
		    elseif clause = eClauseOp.EndWith then
		      Value = kLikeWildcard + CStr(Value)
		    end if
		  end if
		  
		  dim sValue as string
		  sValue = ValueString(Value, AutoQuote)
		  
		  if m_arsWhereClause.Ubound >= 0 then
		    m_arsWhereClause.Append( GetLogicOp(logic) )
		  end if
		  m_arsWhereClause.Append("(")
		  m_arsWhereClause.Append( FieldName(sFieldName,sTableName,AutoBracket) )
		  m_arsWhereClause.Append( GetClauseOp(clause) )
		  m_arsWhereClause.Append( sValue )
		  m_arsWhereClause.Append(")")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddTable(sTableName as string)
		  'Clear if our from type is changing.
		  if m_arsJoin.Ubound >= 0 or m_sComplexFrom <> "" then
		    ClearFromClause
		  end if
		  
		  dim s as string
		  s = Bracket(sTableName, AutoBracket)
		  
		  if m_arsTable.IndexOf(s) < 0 then
		    m_arsTable.Append(s)
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddValue(value as Variant)
		  m_arsValue.Append( ValueString(value, AutoQuote) )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddValues(arvValue() as Variant)
		  'If one of the values won't convert to a string ValueString
		  'throws an exception. That's why we're using a temp. array.
		  dim arsValue() as string
		  for each v as Variant in arvValue
		    arsValue.Append( ValueString(v,AutoQuote) )
		  next
		  
		  for each s as string in arsValue
		    m_arsValue.Append(s)
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub AddValues(value0 as Variant, ParamArray arvRest() as Variant)
		  'Add an array of values.
		  arvRest.Insert(0,value0)
		  AddValues(arvRest)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Bracket(sInputObject as string, bAutoBracket as boolean) As String
		  if not AutoBracket then
		    return sInputObject
		  end if
		  
		  'If it's already got a dot then split it up
		  dim arsDbObject() as string = sInputObject.Split(".")
		  
		  dim i, iMax as integer = arsDbObject.Ubound
		  for i = 0 to iMax
		    dim s as string =arsDbObject(i).Trim
		    if s.InStr(" ")>0 and s.InStr("(")=0 and s.InStr(")")=0 and s.Left(1)<>"`" then
		      arsDbObject(i) = "`" + s + "`"
		    end if
		  next
		  
		  return Join(arsDbObject,".")
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ClearFields()
		  redim m_arsFieldName(-1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ClearFromClause()
		  redim m_arsJoin(-1)
		  redim m_arsTable(-1)
		  m_sComplexFrom = ""
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ClearOrderClause()
		  redim m_arsOrderBy(-1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ClearValues()
		  redim m_arsValue(-1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ClearWhereClause()
		  'redim m_arsWhereClause(-1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor()
		  Reset
		  
		  AutoBracket = true
		  AutoLike = true
		  AutoQuote = true
		  StatementType = eStatementType.Type_Select
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function EvalFrom(byref sFromClause as string) As boolean
		  if m_arsJoin.Ubound >= 0 then
		    sFromClause = MakeJoin
		    return true
		  elseif m_sComplexFrom <> "" then
		    sFromClause = m_sComplexFrom
		    return true
		  elseif m_arsTable.Ubound >= 0 then
		    sFromClause = Join(m_arsTable, ", ")
		    return true
		  else
		    'nothing to select from
		    return false
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FieldName(sField as string, sTableName as string, bAutoBracket as boolean) As string
		  if sTableName = "" then
		    return Bracket(sField, bAutoBracket)
		  else
		    dim sqlTableName as string = Bracket(sTableName,bAutoBracket)
		    dim sqlFieldName as string = Bracket(sField, bAutoBracket)
		    
		    dim x as integer
		    x = sqlFieldName.InStr(".")
		    if x > 0 then
		      'there's already a table name in the field
		      if  sqlFieldName.Left(x-1) = sqlTableName then
		        return sqlFieldName
		      end if
		    end if
		    
		    return sqlTableName + "." + sqlFieldName
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetClauseOp(clause as eClauseOp) As string
		  select case clause
		  case eClauseOp.Equal
		    return "="
		  case eClauseOp.Like, eClauseOp.StartWith, eClauseOp.EndWith
		    return "LIKE"
		  case eClauseOp.GreaterThan
		    return ">"
		  case eClauseOp.GreaterThanOrEqual
		    return ">="
		  case eClauseOp.LessThan
		    return "<"
		  case eClauseOp.LessThanOrEqual
		    return "<="
		  case eClauseOp.NotEqual
		    return "<>"
		  end select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetJoinType(join as eJoinType) As string
		  select case join
		  case eJoinType.Inner
		    return "INNER JOIN"
		  case eJoinType.Left
		    return "LEFT JOIN"
		  case eJoinType.Right
		    return "RIGHT JOIN"
		  case eJoinType.Full
		    return "FULL JOIN"
		  end select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetLogicOp(logic as eWhereClauseLogic) As string
		  select case logic
		  case eWhereClauseLogic.Logic_And
		    return "AND"
		  case eWhereClauseLogic.Logic_Or
		    return "OR"
		  end select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Join(sTable1 as string, sField1 as string, sTable2 as string, sField2 as string, clause as eClauseOp, join as eJoinType)
		  'Join with the current statement
		  
		  if m_arsJoin.Ubound < 0 then
		    redim m_arsTable(-1)
		    m_sComplexFrom = ""
		  end if
		  
		  if m_arsTable.IndexOf(sTable1) < 0 then
		    m_arsTable.Append(sTable1)
		  end if
		  
		  if m_arsTable.IndexOf(sTable2) < 0 then
		    m_arsTable.Append(sTable2)
		  end if
		  
		  'build a string in the form "ON <TABLE1.FIELD1> <OP> <TABLE2.FIELD2>"
		  dim sJoinClause as string
		  sJoinClause = "ON " +  _
		  FieldName(sField1, sTable1, AutoBracket) + _
		  " " + GetClauseOp(clause) + " " + _
		  FieldName(sField2,sTable2, AutoBracket)
		  
		  if m_arsJoin.Ubound = -1 then
		    m_arsJoin.Append( Bracket(sTable1,AutoBracket) )
		  end if
		  m_arsJoin.Append( Bracket(sTable2,AutoBracket) )
		  m_arsJoin.Append( GetJoinType(join)  )
		  m_arsJoin.Append( sJoinClause )
		  m_arsJoin.Append( "." )
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MakeJoin() As string
		  'Convert the join array into a FROM clause.
		  
		  'The join array stores values like so:
		  ' <table1>,<table2>,<join>,<onClause>, "."
		  'The dot tells the us to take the previous four
		  'values and form them into a join.
		  
		  'The way I evaluate this is to run through the join
		  'array and push each value onto a stack, then when
		  'a dot is encounterd pop four values off and build the
		  'join, pushing the result onto the stack.
		  
		  'The advantage of all this is that we can have an unlimited
		  'number of joins by storing something like the following:
		  
		  't1,t2,join1,on1,".",t3,join2,on2,"."
		  
		  dim arsStack() as string
		  for each s as string in m_arsJoin
		    if s = "." then
		      'evaluate the stack contents
		      dim sOnClause as string = arsStack.Pop
		      dim sJoin as string = arsStack.Pop
		      dim sTable2 as string = arsStack.Pop
		      dim sTable1 as string = arsStack.Pop
		      
		      arsStack.Append("(" + sTable1 + " " + sJoin + " "+ sTable2 + " " + sOnClause + ")")
		    else
		      arsStack.Append(s)
		    end if
		  next
		  
		  return arsStack(0)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MakeStatement() As string
		  'Build the SQL statement.
		  
		  'Pieces of the statement are added to arsSQL and at the end they are joined
		  'together and returned.
		  
		  dim arsSQL() as string
		  
		  dim sBaseTable as string
		  if m_arsTable.Ubound >= 0 then
		    sBaseTable = m_arsTable(0)
		  end if
		  
		  if StatementType = eStatementType.Type_Select then
		    dim sFromClause as string
		    if EvalFrom(sFromClause) then
		      arsSQL.Append("SELECT")
		      if m_arsFieldName.Ubound < 0 then
		        arsSQL.Append("*")
		      else
		        arsSQL.Append( Join(m_arsFieldName, ",") )
		      end if
		      arsSQL.Append("FROM")
		      arsSQL.Append(sFromClause)
		    end if
		  elseif StatementType = eStatementType.Type_Insert then
		    if sBaseTable.Len > 0 and m_arsValue.Ubound >= 0 then
		      arsSQL.Append("INSERT INTO")
		      arsSQL.Append(sBaseTable)
		      if m_arsFieldName.Ubound >= 0 then
		        arsSQL.Append("(" + Join(m_arsFieldName,",") + ")")
		      end if
		      arsSQL.Append("VALUES")
		      arsSQL.Append("(" + Join(m_arsValue,",") + ")")
		    end if
		  elseif StatementType = eStatementType.Type_Update then
		    if sBaseTable.Len > 0 then
		      dim arsValues() as string
		      dim i,iMax as integer
		      iMax = m_arsFieldName.Ubound
		      for i = 0 to iMax
		        if i > m_arsValue.Ubound then
		          exit for
		        end if
		        arsValues.Append(m_arsFieldName(i) + "=" + m_arsValue(i))
		      next
		      
		      if arsValues.Ubound >= 0 then
		        arsSQL.Append("UPDATE")
		        arsSQL.Append(sBaseTable)
		        arsSQL.Append("SET")
		        arsSQL.Append(Join(arsValues,","))
		      end if
		    end if
		  elseif StatementType = eStatementType.Type_Delete then
		    if sBaseTable.Len > 0 then
		      arsSQL.Append("DELETE FROM")
		      arsSQL.Append(sBaseTable)
		    end if
		  end if
		  
		  if arsSQL.Ubound < 0 then
		    'there's no command
		    return ""
		  end if
		  
		  'WHERE clause
		  if StatementType <> eStatementType.Type_Insert  and m_arsWhereClause.Ubound>=0 then
		    arsSQL.Append("WHERE")
		    arsSQL.Append( Join(m_arsWhereClause," ") )
		  end if
		  
		  'ORDER BY clause
		  if StatementType = eStatementType.Type_Select and m_arsOrderBy.Ubound>=0 then
		    arsSQL.Append("ORDER BY")
		    arsSQL.Append( Join(m_arsOrderBy, ",") )
		  end if
		  
		  return Join( arsSQL, " " ) + ";"
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Quote(s as string, bAutoQuote as boolean) As string
		  'Replace single quote (') with two single quotes
		  'Enclose the string in single quotes
		  if bAutoQuote then
		    return "'" + ReplaceAll(s, "'", "''") + "'"
		  else
		    return s
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Reset()
		  ClearFromClause
		  ClearWhereClause
		  ClearOrderClause
		  ClearFields
		  ClearValues
		  StatementType = eStatementType.Type_Select
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetupJoin(sTable1 as string, sField1 as string, sTable2 as string, sField2 as string)
		  SetupJoin(sTable1,sField1,sTable2,sField2,eClauseOp.Equal,eJoinType.Inner)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetupJoin(sTable1 as string, sField1 as string, sTable2 as string, sField2 as string, clause as eClauseOp)
		  SetupJoin(sTable1,sField1,sTable2,sField2,clause, eJoinType.Inner)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetupJoin(sTable1 as string, sField1 as string, sTable2 as string, sField2 as string, clause as eClauseOp, join as eJoinType)
		  ClearFromClause
		  
		  Join(sTable1,sField1,sTable2,sField2, clause, join)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ValueString(value as Variant, bAutoQuote as boolean) As string
		  'Try to convert a value to a string. We'll accept anything that can convert.
		  if value = nil then
		    return "NULL"
		  end if
		  
		  select case VarType(value)
		  case Variant.TypeDouble, Variant.TypeSingle
		    'if it can be represented as a 64bit integer, then format it as an integer
		    if value.Int64Value = value.DoubleValue then
		      return Str( value.Int64Value )
		    else
		      return Str( value.DoubleValue )
		    end if
		  case Variant.TypeBoolean
		    if value then
		      return kTrueValue
		    else
		      return kFalseValue
		    end if
		  case Variant.TypeObject
		    if value isa Date then
		      return Date(value).SQLDateTime
		    else
		      return Quote(CStr(value),bAutoQuote)
		    end if
		  case else
		    dim sValue as string = CStr(value)
		    if IsNumeric(sValue) and VarType(value) <> Variant.TypeString then
		      return sValue
		    end if
		    return Quote(sValue, bAutoQuote)
		  end select
		  
		exception err as TypeMismatchException
		  'the value couldn't be converted to a string
		  raise new cInvalidValueException
		End Function
	#tag EndMethod


	#tag Property, Flags = &h0
		AutoBracket As boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		AutoLike As boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		AutoQuote As boolean
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return m_sComplexFrom
			End Get
		#tag EndGetter
		#tag Setter
			Set
			  ClearFromClause
			  
			  dim sFrom as string = value.Trim
			  if sFrom.Left(4) = "FROM" then
			    sFrom = sFrom.Mid(5).Trim
			  end if
			  
			  'parse out the table names
			  dim arsCommaSep() as string = sFrom.Split(",")
			  for each sPart as string in arsCommaSep
			    dim arsJoinSep() as string = sPart.Split("JOIN")
			    
			    dim sTableName as string
			    
			    'the first item of each should be the table
			    for each sTable as string in arsJoinSep
			      sTableName = ReplaceAll(sTable,"(","").Trim
			      dim x as integer = sTableName.InStr(" ")
			      if x > 0 then
			        sTableName = sTableName.Left(x-1)
			      end if
			      if m_arsTable.IndexOf(sTableName) < 0 then
			        m_arsTable.Append(sTableName)
			      end if
			    next
			  next
			  
			  m_sComplexFrom = sFrom
			End Set
		#tag EndSetter
		ComplexFromClause As String
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private m_arsFieldName() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m_arsJoin() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m_arsOrderBy() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m_arsTable() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m_arsValue() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m_arsWhereClause() As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private m_sComplexFrom As string
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return MakeStatement
			End Get
		#tag EndGetter
		SQL As string
	#tag EndComputedProperty

	#tag Property, Flags = &h0
		StatementType As eStatementType
	#tag EndProperty

	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return m_arsTable.Ubound + 1
			End Get
		#tag EndGetter
		TableCount As Integer
	#tag EndComputedProperty


	#tag Constant, Name = kFalseValue, Type = String, Dynamic = False, Default = \"0", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kLikeWildcard, Type = String, Dynamic = False, Default = \"%", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kTrueValue, Type = String, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="AutoBracket"
			Group="Behavior"
			InitialValue="0"
			Type="boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AutoLike"
			Group="Behavior"
			InitialValue="0"
			Type="boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AutoQuote"
			Group="Behavior"
			InitialValue="0"
			Type="boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ComplexFromClause"
			Group="Behavior"
			Type="String"
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
			Name="SQL"
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
			Name="TableCount"
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
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
