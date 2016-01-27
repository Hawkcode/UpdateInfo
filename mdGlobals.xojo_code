#tag Module
Protected Module mdGlobals
	#tag Method, Flags = &h0
		Sub Assertion(Condition as Boolean, msg as string)
		  if not condition then
		    msgbox "Assertion failed :" + msg
		    try
		      raise new NilObjectException
		    end try
		    
		    break
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Busy(isBusy as Boolean, prgWheel as WebProgressWheel)
		  if isBusy then
		    'prgWheel.Visible = True
		    'app.MouseCursor =  system.Cursors.Wait
		    'else
		    'prgWheel.Visible = False
		    'app.MouseCursor =  system.Cursors.StandardPointer
		  end
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CheckDate(lsDate as DatabaseCursorField) As String
		  
		  if lsDate.StringValue <> "" and  lsDate <> nil and  lsDate.StringValue <> "0-00-00 00:00:00"  then
		    Return lsDate.DateValue.ShortDate
		    
		  end
		  
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CheckDateText(lsDate as String) As Date
		  Dim ldDate as Date
		  
		  ldDate = StrToDate(lsDate)
		  if lsDate <> ""  then
		    Return ldDate
		    
		  end
		  
		  Return nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ChkStr(lsStr as Variant) As String
		  
		  if lsStr = nil then
		    return ""
		  else
		    return lsStr
		  end
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CopyFolder(FoldernameFrom as folderItem, FoldernameTo as folderitem)
		  
		  Dim i as integer
		  if FoldernameFrom.Directory = false then return
		  FoldernameTo.CreateAsFolder
		  for i = 1 to FoldernameFrom.Count
		    if FoldernameFrom.item(i).Directory then
		      dim subfolder as FolderItem = GetFolderItem(FoldernameTo.AbsolutePath + FoldernameFrom.Item(i).Name)
		      CopyFolder (FoldernameFrom.item(i),subfolder)
		    else
		      dim subfile as FolderItem = GetFolderItem(FoldernameTo.AbsolutePath + FoldernameFrom.Item(i).Name)
		      FoldernameFrom.Item(i).CopyFileTo (subfile)
		    end if
		  next
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CreatePrintPreferences()
		  Dim BinStream as BinaryStream
		  Dim theString as String
		  Dim Prefs as FolderItem
		  
		  Prefs = SpecialFolder.Preferences.Child("Printing Test Print Settings")
		  
		  if Prefs <> nil then
		    BinStream = Prefs.CreateBinaryFile("myPrefsFile")
		    BinStream.Write gPrintSettings
		    BinStream.close
		  end
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DateSql(lv as String) As Variant
		  dim d as Date
		  if lv = "" then
		    return nil
		  end
		  
		  If ParseDate( lv, d) then
		    
		    
		    return d.SQLDate
		    
		  Else
		    return Nil
		  End
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DatetoUnix(d As Date) As Int64
		  Return (d.TotalSeconds - 2082844800)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function DeleteEntireFolder(theFolder as FolderItem, continueIfErrors as Boolean = false, lbContentsOnly as Boolean = true) As Integer
		  // Returns an error code if it fails, or zero if the folder was deleted successfully
		  
		  dim returnCode, lastErr, itemCount as integer
		  dim files(), dirs() as FolderItem
		  
		  if theFolder = nil or not theFolder.Exists() then
		    return 0
		  end if
		  
		  // Collect the folder‘s contents first.
		  // This is faster than collecting them in reverse order and deleting them right away!
		  itemCount = theFolder.Count
		  for i as integer = 1 to itemCount
		    dim f as FolderItem
		    f = theFolder.TrueItem( i )
		    if f <> nil then
		      if f.Directory then
		        dirs.Append f
		      else
		        files.Append f
		      end if
		    end if
		  next
		  
		  // Now delete the files
		  for each f as FolderItem in files
		    f.Delete
		    lastErr = f.LastErrorCode   // Check if an error occurred
		    if lastErr <> 0 then
		      if continueIfErrors then
		        if returnCode = 0 then returnCode = lastErr
		      else
		        // Return the error code if any. This will cancel the deletion.
		        return lastErr
		      end if
		    end if
		  next
		  
		  redim files(-1) // free the memory used by the files array before we enter recursion
		  
		  // Now delete the directories
		  for each f as FolderItem in dirs
		    lastErr = DeleteEntireFolder( f, continueIfErrors )
		    if lastErr <> 0 then
		      if continueIfErrors then
		        if returnCode = 0 then returnCode = lastErr
		      else
		        // Return the error code if any. This will cancel the deletion.
		        return lastErr
		      end if
		    end if
		  next
		  
		  If not lbContentsOnly then
		    if returnCode = 0 then
		      // We‘re done without error, so the folder should be empty and we can delete it.
		      theFolder.Delete
		      returnCode = theFolder.LastErrorCode
		    end if
		  End if
		  return returnCode
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DoCSVFromRS(rs as RecordSet, lfFile as FolderItem, lsSql as String, launch as Boolean = false)
		  dim tos as textOutputStream
		  dim s as string
		  dim i, j as integer
		  
		  'show standard file selector
		  
		  'create file
		  'tos = f.CreateTextFile
		  'if tos = nil then                                       'failed?
		  'MsgBox("The file could not be created!")
		  'exit sub
		  'end if
		  
		  if lfFile.exists then
		    lfFile.Delete
		    
		    'if f.LastErrorCode > 0 then
		    'MsgBox("Unable to overwrite excel file: "+lfFile.Name)
		    'return
		    'end
		  end
		  
		  tos = lfFile.CreateTextFile
		  'tos.Delimiter=Chr(10)  'LineFeed
		  if tos = nil then                                       'failed?
		    MsgBox("The file could not be created!")
		    exit sub
		  end if
		  
		  // walk over all fields
		  s=""""
		  
		  for i  = 1 to rs.FieldCount
		    s = s +rs.IdxField(i).Name + ""","""
		  next
		  tos.WriteLine(s.left(s.len-2))
		  
		  for i=1 to rs.RecordCount            'for each row
		    s=""""                                                   'build line to save
		    for j=1 to rs.FieldCount   'for each column
		      s = s +ChkStr(rs.IdxField(j).StringValue) + ""","""
		    next
		    tos.WriteLine(s.left(s.len-2))
		    rs.MoveNext              'save line
		  next
		  
		  tos.Close
		  
		  
		  
		  
		  
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function formatDateTime(d as date, mask as string) As String
		  // mask like "MM-DD-YYYY HH:NN:SS" OR or "M-D-YY H:N:S"
		  //Note Minute = "N" not "M" (which is month)
		  
		  Dim s As String
		  Dim i As Integer
		  
		  s = format(d.Year,"0000")
		  mask = replace(mask,"YYYY",s)
		  s = Right(s,2)
		  mask = replace(mask,"YY",s)
		  s = format(d.Year,"0")
		  mask = replace(mask,"Y",s)
		  
		  s = format(d.Month,"00")
		  mask = replace(mask,"MM",s)
		  s = format(d.Month,"0")
		  mask = replace(mask,"M",s)
		  
		  s = format(d.Day,"00")
		  mask = replace(mask,"DD",s)
		  s = format(d.Day,"0")
		  mask = replace(mask,"D",s)
		  
		  s = format (d.Hour,"00")
		  mask = replace(mask,"HH",s)
		  s = format (d.Hour,"0")
		  mask = replace(mask,"H",s)
		  
		  s = format( d.Minute,"00")
		  mask = replace(mask,"NN",s)
		  s = format( d.Minute,"0")
		  mask = replace(mask,"N",s)
		  
		  s = format( d.Second,"00")
		  mask = replace(mask,"SS",s)
		  s = format( d.Second,"0")
		  mask = replace(mask,"S",s)
		  
		  Return mask
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetPID(lnUID as Integer = 0) As Integer
		  DIM lsSql, lsEmail AS STRING
		  Dim rsD as RecordSet
		  Dim lbFound as Boolean
		  
		  if not Session.Available then
		    app.WriteLog("Session Not Available: ")
		  else
		    app.WriteLog("Session is Available: ")
		  end
		  
		  
		  Try
		    lbFound = False
		    if lnUID <> 0 then
		      lsSql = "SELECT d_profile_value.fid, d_profile_value.value, d_profile_value.uid FROM d_profile_value "
		      lsSql = lsSql + "WHERE d_profile_value.fid = 17 and d_profile_value.uid = " + Str(lnUID)
		      
		      'MsgBox(lssql)
		      'if not ConnectWS then
		      '
		      'Return 0
		      'end
		      '
		      rsD = Session.sesWebDB.SQLSelect(lsSql)
		      if  not IsNull(rsD) then
		        if not Session.sesWebDB.CheckDBError then
		          if rsD.RecordCount  > 0 then
		            Return rsD.Field("value").IntegerValue
		          end
		        end
		      end
		    end
		  Catch e As RunTimeException
		    app.WriteLog("Using UID an Exception of type: " + e.Type + " Message: " + e.Message )
		    Return 0
		  End Try
		  
		  if lnUID = 0 then Return 0
		  
		  lsSql = "Select mail from d_users where uid = " + Str(lnUID)
		  
		  rsD = Session.sesWebDB.SQLSelect(lsSql)
		  if not Session.sesWebDB.CheckDBError then
		    Try
		      if rsD.RecordCount  > 0 then
		        
		        lsEmail = rsD.Field("value").StringValue
		      else
		        Return 0
		      end
		      
		      lsSql = "Select PersonID from tblPeople where Email = '" + lsEmail + "' "
		      rsD = Session.sesAspeDB.SQLSelect(lsSql)
		    Catch e As RunTimeException
		      app.WriteLog("Using Email an Exception of type: " + e.Type + EndOfLine + _
		      "                           Reason: " + e.Reason + EndOfLine + _
		      "                     Error Number: " + e.ErrorNumber.ToText + EndOfLine + _
		      "                          Message: " + e.Message + EndOfLine +  _
		      "SQL: >>" + lsSql + "<<" + EndOfLine + EndOfLine )
		      Return 0           '"                            Stack: " + e.Stack() + _
		      
		      
		    end try
		    Try
		      
		      if  not IsNull(rsD) then
		        if not Session.sesAspeDB.CheckDBError then
		          if rsD.RecordCount  > 0 then
		            Return rsD.Field("PersonID").IntegerValue
		          else
		            Return 0
		          end
		        end
		        
		      end
		    Catch e As RunTimeException
		      app.WriteLog("Returning PID an Exception of type: " + e.Type + " Message: " + e.Message )
		      Return 0
		      
		    end try
		    
		  end
		  
		  return 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function InstrTrue(start as integer = 0, stringtosearch as String, stringtofind as String) As Boolean
		  
		  // returns a Boolean because rb insists on returning an integer
		  //this way you can write "if instrtrue("Boolean","boo") if all you want to know is:
		  //does "Boolean" contain "boo"
		  
		  dim a as integer
		  a = instr(start,stringtosearch,stringtofind)
		  if a > 0 then
		    return true
		  else
		    return false
		  end
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsNotValid(Win as WebContainer) As Boolean
		  Dim lnX, lnCount As Integer
		  
		  lnCount = Win.ControlCount - 1
		  
		  
		  for lnX = 0 to lnCount
		    if Win.ControlAtIndex(lnX).Name = "txtFirst" then 'EntryFieldsError then
		      Return True
		    End
		    
		    'if Win.Control(lnX) IsA TextField then
		    'TextField(Win.Control(lnX)).Text = ""
		    'end
		    'if Win.Control(lnX) IsA CheckBox then
		    'CheckBox(Win.Control(lnX)).Value = False
		    'end
		    'if Win.Control(lnX) IsA ComboBox then
		    'ComboBox(Win.Control(lnX)).ListIndex = -1
		    'end
		    'if Win.Control(lnX) IsA PopupMenu then
		    'PopupMenu(Win.Control(lnX)).ListIndex = -1
		    'end
		    
		  next
		  Return False
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadListBox(rs as RecordSet, lstResults as WebListBox)
		  Dim lsSql, lsHeader as String
		  Dim lnCnt, i, lnColType, lnTemp as Integer
		  Dim lsStr, lsColWidths, msHeaderRow as String
		  
		  
		  'lstResults.Heading InitialValue = ""
		  lstResults.DeleteAllRows
		  'lblResult.Text = Str(rs.RecordCount )
		  
		  
		  grsHold= rs
		  
		  
		  
		  // query number of fields
		  lnCnt = rs.FieldCount
		  lstResults.ColumnCount = lnCnt + 1
		  // walk over all fields
		  
		  msHeaderRow = ""
		  lsHeader = "Dummy" + Chr(9)
		  for i  = 1 to lnCnt
		    lsHeader = lsHeader +rs.IdxField(i).Name + Chr(9)
		    msHeaderRow= msHeaderRow + rs.IdxField(i).Name + ","
		  next
		  'lstResults.InitialValue =lsHeader
		  
		  lsColWidths = "0,"
		  for i  = 1 to lnCnt
		    lsStr = rs.IdxField(i).StringValue
		    lnColType = rs.ColumnType(i-1)
		    Select case lnColType
		      
		    Case 0, 1, 2, 3   'Int
		      lsColWidths = lsColWidths + "70,"
		    Case 4 'Char
		      lsColWidths = lsColWidths + "10,"
		    Case 5 'varchar
		      lsStr = rs.IdxField(i).StringValue
		      if lsStr.Len <= 3 then
		        lsColWidths = lsColWidths + "50,"
		      elseif lsStr.Len <= 10 then
		        lsColWidths = lsColWidths + "200,"
		      ElseIf lsStr.Len <= 25 then
		        lsColWidths = lsColWidths + "300,"
		      else
		        lsColWidths = lsColWidths + "400,"
		      end
		      
		    Case 6, 7  'Float,Double
		      lsColWidths = lsColWidths + "80,"
		      'lstResults.ColumnAlignment(i-1)= Listbox.AlignRight
		      
		    Case 8   'Date
		      lsColWidths = lsColWidths + "90,"
		      
		    case 9, 10 'Time, Timestamp
		      lsColWidths = lsColWidths + "90,"
		      
		    Case 11 ' Currency
		      lsColWidths = lsColWidths + "100,"
		      'lstResults.ColumnAlignment(i-1)= Listbox.AlignRight
		      
		    Case 12 'Boolean
		      lsColWidths = lsColWidths + "20"
		      'lstResults.ColumnType(i-1)=ListBox.TypeCheckbox
		      
		    case 13 'Decimal
		      lsColWidths = lsColWidths + "100,"
		      'lstResults.ColumnAlignment(i-1)= Listbox.AlignRight
		      
		    Case 14 'Binary
		      lsColWidths = lsColWidths + "100,"
		      'lstResults.ColumnAlignment(i-1)= Listbox.AlignRight
		      
		    Case 15 'Long Text Blob
		      lsColWidths = lsColWidths + "300,"
		      
		      'Case 16 'Long VarBinary
		      'Case 17 'MacPict
		    Case 18 ' String same as VarChr
		      lsColWidths = lsColWidths + "200,"
		      
		    Case 19 ' int64
		      lsColWidths = lsColWidths + "100,"
		      'lstResults.ColumnAlignment(i-1)= Listbox.AlignRight
		      
		    else
		      lsStr = rs.IdxField(i).StringValue
		      lnTemp = instr(0, lsStr,".")
		      
		      if (lnTemp + 2) = Len(lsStr) then
		        lsColWidths = lsColWidths + "80,"
		        'lstResults.ColumnAlignment(i-1)= Listbox.AlignRight
		        
		      else
		        lsColWidths = lsColWidths + "100,"
		      end
		    end select
		  next
		  lstResults.ColumnWidths = lsColWidths
		  'lstResults.ColumnsResizable = True
		  
		  
		  
		  While (Not rs.EOF)
		    for i  = 1 to lnCnt +1
		      if i = 1 then
		        
		        lstResults.AddRow ""
		      else
		        if rs.ColumnType(i-2) >= 8  and rs.ColumnType(i-2) <= 10 then 'Date
		          if rs.IdxField(i-1).DateValue = Nil then
		            lsStr = ""
		          else
		            lsStr = rs.IdxField(i-1).DateValue.ShortDate
		          end
		          lstResults.Cell(lstResults.LastIndex, i-1) = lsStr
		        elseif rs.ColumnType(i-2) = 7 then
		          lsStr = Format(rs.IdxField(i-1).CurrencyValue, "###.00")
		          'lsStr = rs.IdxField(i).StringValue
		          lstResults.Cell(lstResults.LastIndex, i-1) = lsStr
		          
		        else
		          
		          lsStr = rs.IdxField(i-1).StringValue
		          lstResults.Cell(lstResults.LastIndex, i-1) = lsStr
		        end
		      end
		    Next
		    rs.MoveNext
		  wend
		  
		  'lstResults.ColumnCount = 12
		  'lstResults.InitialValue = "ProductID" + Chr(9) + "PersonID" + Chr(9) + "OrderID" + Chr(9) + "OrderDate" + Chr(9) + "Description" + Chr(9) + "Qty" + Chr(9) + _
		  '"Price" + Chr(9) + "ItemTotal" + Chr(9) + "Account" + Chr(9) + "Category" + Chr(9) + "State" + Chr(9) + "Country"
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LoadListBoxBasic(rs as RecordSet, lstResults as WebListbox)
		  Dim lsSql, lsHeader as String
		  Dim lnCnt, i, lnColType, lnTemp as Integer
		  Dim lsStr, lsColWidths, msHeaderRow as String
		  
		  
		  'lstResults.InitialValue = ""
		  lstResults.DeleteAllRows
		  'lblResult.Text = Str(rs.RecordCount )
		  
		  
		  grsHold= rs
		  
		  lnCnt = rs.FieldCount
		  
		  // query number of fields
		  
		  
		  While (Not rs.EOF)
		    for i  = 1 to lnCnt
		      lsStr = rs.IdxField(i).StringValue
		      if i = 1 then
		        lstResults.AddRow lsStr
		      else
		        if rs.ColumnType(i-1) >= 8  and rs.ColumnType(i-1) <= 10 then 'Date
		          if rs.IdxField(i).DateValue = Nil then
		            lsStr = ""
		          else
		            lsStr = rs.IdxField(i).DateValue.ShortDate
		          end
		          lstResults.Cell(lstResults.LastIndex, i-1) = lsStr
		        elseif rs.ColumnType(i) = 7 then
		          lsStr = Format(rs.IdxField(i).CurrencyValue, "###.00")
		          'lsStr = rs.IdxField(i).StringValue
		          lstResults.Cell(lstResults.LastIndex, i-1) = lsStr
		          
		        else
		          
		          lsStr = rs.IdxField(i).StringValue
		          lstResults.Cell(lstResults.LastIndex, i-1) = lsStr
		        end
		      end
		    Next
		    rs.MoveNext
		  wend
		  
		  'lstResults.ColumnCount = 12
		  'lstResults.InitialValue = "ProductID" + Chr(9) + "PersonID" + Chr(9) + "OrderID" + Chr(9) + "OrderDate" + Chr(9) + "Description" + Chr(9) + "Qty" + Chr(9) + _
		  '"Price" + Chr(9) + "ItemTotal" + Chr(9) + "Account" + Chr(9) + "Category" + Chr(9) + "State" + Chr(9) + "Country"
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function MonthName(month As Integer, LongName as Boolean) As String
		  if (month >= 1) and (month <= 12) then
		    if LongName then
		      return NthField("January,February,March,April,May,June,July,August,September,October,November,December",",",month)
		    else
		      return NthField("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec",",",month)
		    end
		  else
		    return "" // Invalid month number - maybe better to raise an exception
		  end if
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub MoveFolderContents(lfiSource as folderItem, lfiTarget as folderitem)
		  Dim lnCnt , itemCount as Integer
		  
		  
		  
		  if lfiSource.Directory = false then return
		  
		  itemCount = lfiSource.Count
		  for lnCnt = itemCount DownTo 1
		    'if lnCnt = 144 then Break
		    
		    dim f as FolderItem
		    f = lfiSource.TrueItem( lnCnt )
		    if f <> nil then
		      if Not f.Directory then
		        f.MoveFileTo(lfiTarget)
		      end if
		    end if
		  Next
		  
		  'lfiSource.Delete
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub OpenPrintPreferences()
		  Dim BinStream as BinaryStream
		  Dim Prefs as FolderItem
		  
		  Prefs = SpecialFolder.Preferences.Child("Printing Test Print Settings")
		  
		  if Prefs.Exists then
		    binStream = Prefs.OpenAsBinaryFile(true)
		    gPrintSettings = BinStream.Read(BinStream.length +1)
		    BinStream.close
		  else
		    CreatePrintPreferences
		  end
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ParseXML(KeyQuery as String, Data as String) As String
		  Dim lsKeyStart as String = "<" + KeyQuery + ">"
		  Dim lsKeyEnd as String = "</" + KeyQuery + ">"
		  
		  Return Mid(Data, InStr(Data, lsKeyStart) + Len(lsKeyStart), InStr(Data, lsKeyEnd) - InStr(Data, lsKeyStart) + Len(lsKeyStart))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetEnableAllTextFields(Win as WebContainer, lbEnable as Boolean)
		  Dim lnX, lnCount As Integer
		  
		  lnCount = Win.ControlCount - 1
		  
		  
		  for lnX = 0 to lnCount
		    if Win.ControlAtIndex(lnX) IsA WebTextField then
		      WebTextField(Win.ControlAtIndex(lnX)).Enabled = lbEnable
		    end
		    if Win.ControlAtIndex(lnX) IsA WebCheckBox then
		      WebCheckBox(Win.ControlAtIndex(lnX)).Enabled = lbEnable
		    end
		    if Win.ControlAtIndex(lnX) IsA WebPopupMenu then
		      WebPopupMenu(Win.ControlAtIndex(lnX)).Enabled = lbEnable
		    end
		    
		  next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SetPopIndex(pop As WebPopupMenu, lsSearch As String) As Integer
		  
		  
		  Dim lnCnt as Integer
		  
		  
		  For lnCnt = 0 to pop.ListCount - 1
		    pop.ListIndex = lnCnt
		    if pop.Text = lsSearch then
		      Return lnCnt
		    end
		    
		    
		  next
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetPopMenuValue(Extends theMenu as WebPopupMenu, theValue as String, lnCharFromLeft as Integer = 0)
		  dim i as Integer
		  
		  If theMenu Is Nil then
		    Return
		  End if
		  
		  For i = 0 to theMenu.ListCount - 1
		    if lnCharFromLeft = 0 then
		      If theMenu.List(i) = theValue then
		        theMenu.ListIndex = i
		        Return
		      End if
		    else
		      If theMenu.List(i).Left(lnCharFromLeft) = theValue then
		        theMenu.ListIndex = i
		        Return
		      end
		    end
		  Next
		  theMenu.ListIndex = -1
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Str2Int(s as String) As Integer
		  
		  Return CType(Val(s), Integer)
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function StrToDate(lStr as String) As Date
		  DIm ld As  Date
		  
		  if ParseDate(lStr, ld) then
		    Return ld
		  else
		    Return ToDay
		  end
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Today() As Date
		  Dim dtNow as Date
		  dtNow = New Date
		  Return dtNow
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Type(extends e As RuntimeException) As String
		  Dim t As Introspection.TypeInfo = Introspection.GetType(e)
		  If t <> Nil Then
		    Return t.FullName
		  Else
		    //this should never happen...
		    Return ""
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function UnixtoDate(i As Int64) As Date
		  Dim cd As New Date
		  
		  i = i + 2082844800
		  cd.TotalSeconds = i
		  
		  Return cd
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ValidateEmail(lsEmail as String) As Boolean
		  
		  Dim myRegEx As RegEx
		  Dim myMatch As RegExMatch
		  myRegEx = New RegEx
		  myRegEx.Options.TreatTargetAsOneLine = True
		  myRegEx.SearchPattern = "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
		  ' If no Match Valid Email myRegEx.SearchPattern ="[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
		  'Pop up all matches one by one
		  myMatch = myRegEx.Search(lsEmail)
		  If myMatch <> Nil then
		    'MsgBox("OK")
		    Return True
		  else
		    Return False
		  end
		  
		  
		  '"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function validDate(text as String, ByRef value As Date, assumePastFuture as integer = 0) As Boolean
		  // If the year provided has only one or 2 digits, check assumePastFuture
		  // negative value means the past, positive means future, 0 means current century
		  
		  // If no year is supplied, assumePastFuture has a granularity of 1 year
		  
		  Static yearPos as Integer = -9
		  Static monthPos as Integer = -9
		  Static dayPos as Integer = -9
		  
		  if yearPos = -9 then // first time through
		    yearPos = -1 // only try this once
		    // try to work out local date format
		    // assume Gregorian calendar
		    // assume shortDate contains all 3 numbers
		    // don't use NthField or Split in case it contains other characters
		    dim d as new date
		    // clear any time numbers, in case of unusual shortDate format
		    d.TotalSeconds = 0
		    // set unique values for year, month & date
		    d.SQLDate = "2005-12-31"
		    dim s as String = d.ShortDate
		    dim pos() As integer
		    dim thisPos  as Integer
		    
		    thisPos = InStr( 0, s, "05" )
		    if thisPos > 0  then pos.Append thisPos
		    thisPos = InStr( 0, s, "12" )
		    if thisPos > 0  then pos.Append thisPos
		    thisPos = InStr( 0, s, "31" )
		    if thisPos > 0  then pos.Append thisPos
		    
		    if UBound( pos ) = 2 then
		      // we've found all three elements
		      // sort them by position in shortDate
		      dim typ() As string = Array( "y", "m", "d" )
		      pos.SortWith typ
		      yearPos = typ.IndexOf( "y" )
		      monthPos = typ.IndexOf( "m" )
		      dayPos = typ.IndexOf( "d" )
		    end if
		    
		  end if
		  if yearPos < 0 or monthPos < 0 or dayPos < 0 then
		    // we don't know how to parse the date
		    // some might want to set defaults instead of returning false
		    Return false
		  end if
		  
		  // now check the date has just numbers and two delimiters
		  Dim sep as String
		  Dim tmp as String
		  Dim i as Integer
		  Dim noYearSupplied as Boolean
		  
		  tmp = text
		  
		  // first figure out what separator they gave us .. have to both be the same one
		  for i = 0 to 9
		    tmp = replaceAll(tmp,format(i,"0"),"")
		  next
		  
		  select case len(tmp)
		  case 0
		    // unable to understand the format entered
		    return false
		  case 1
		    sep = tmp
		  case 2
		    
		    sep = mid(tmp,1,1)
		    if sep <> mid(tmp,2,1) then
		      // invalid - two different separators
		      return false
		    end if
		    
		  else
		    return false
		  end select
		  
		  //make array of elements
		  Dim dats() as String = Split( text, sep )
		  
		  if UBound( dats ) <> 2 then
		    // add in the missing year ?
		    dim tmpDate as new date
		    dats.Insert  yearPos, format(tmpDate.year,"0000")
		    noYearSupplied = True
		  end if
		  
		  if UBound( dats ) <> 2 then
		    //invalid date - should never get here.
		    Return false
		  end if
		  
		  dim yr As integer = CDbl( dats( yearPos ) )
		  if yr < 100 then
		    // fix short year by assuming current century
		    // proving that we learned nothing from y2k
		    dim today as new Date
		    dim century as integer
		    century = today.year \ 100
		    century = century * 100
		    yr = yr + century
		    // use any assumptions about whether the date is past or future to set century
		    if assumePastFuture < 0 then
		      if yr > today.Year then
		        yr = yr - 100
		      end if
		    elseif assumePastFuture > 0 and yr < today.Year then
		      yr = yr + 100
		    end if
		    dats( yearPos ) = CStr( yr )
		  elseif noYearSupplied then
		    // use any assumptions about whether the date is past or future to set year
		    dim mth as integer = CDbl( dats( monthPos ) )
		    dim dy as integer = CDbl( dats( dayPos ) )
		    dim today as new Date
		    if assumePastFuture < 0 then
		      if mth > today.Month or ( mth = today.Month and dy > today.Day )  then
		        yr = yr - 1
		      end if
		    elseif assumePastFuture > 0 and ( mth < today.Month or ( mth = today.Month and dy < today.Day ) ) then
		      yr = yr + 1
		    end if
		    dats( yearPos ) = CStr( yr )
		  end if
		  
		  // put detail into a date object
		  dim retVal as new date
		  dim yy,mm,dd as Integer
		  yy = val( dats( yearPos ) )
		  mm = val( dats( monthPos ) )
		  dd = val( dats( dayPos ) )
		  
		  retVal.TotalSeconds = 0
		  retVal.Year = yy
		  retVal.Month = mm
		  retVal.Day = dd
		  
		  // check the date object is not making corrections
		  if retVal.Year <> yy or retVal.Month <> mm or retVal.Day <> dd then
		    //probably an invalid day of the month
		    Return false
		  end if
		  
		  //populate value ( ByRef side-effect )
		  if value = nil Then
		    value = new Date
		  end if
		  value.totalseconds = retVal.TotalSeconds
		  return true
		  
		End Function
	#tag EndMethod


	#tag Note, Name = Left off
		Testing creating mbrshp and renewing.
		
		#Pragma BreakOnExceptions Off use to turn off in one function.
		
		
		Need to make chapters cbo load from table on People Screen
		
		Membership New, OK
		
		Membership Renew, OK
		
		Membership Old Re-instate OK
		
		Membership Old to new OK
	#tag EndNote


	#tag Property, Flags = &h0
		gbClose As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		gbGroupSent As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		gbLookUpPID As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		gbOrderIsARegistration As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		gbPostPayment As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		gbRegistrationActive As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h0
		gCurrentUser As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gdAmount As Currency
	#tag EndProperty

	#tag Property, Flags = &h0
		gdPayment As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		gdTotalDue As Double
	#tag EndProperty

	#tag Property, Flags = &h0
		gfiEOMArchivePath As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		gfiEOMPath As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		gfiReportsFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		gfiSavePDFPath As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		gnEmailPort As Integer = 465
	#tag EndProperty

	#tag Property, Flags = &h0
		gnFinanceHeight As Integer = 600
	#tag EndProperty

	#tag Property, Flags = &h0
		gnFinanceLeft As Integer = 20
	#tag EndProperty

	#tag Property, Flags = &h0
		gnFInanceTop As Integer = 50
	#tag EndProperty

	#tag Property, Flags = &h0
		gnFinanceWidth As Integer = 980
	#tag EndProperty

	#tag Property, Flags = &h0
		gnHomeLeft As Integer = 10
	#tag EndProperty

	#tag Property, Flags = &h0
		gnHomeTop As Integer = 100
	#tag EndProperty

	#tag Property, Flags = &h0
		gnOrderID As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		gnOrderLeft As Integer = 200
	#tag EndProperty

	#tag Property, Flags = &h0
		gnOrderTop As Integer = 200
	#tag EndProperty

	#tag Property, Flags = &h0
		gnPeopleLeft As Integer = 200
	#tag EndProperty

	#tag Property, Flags = &h0
		gnPeopleTop As Integer = 200
	#tag EndProperty

	#tag Property, Flags = &h0
		gnRegistrationLeft As Integer = 10
	#tag EndProperty

	#tag Property, Flags = &h0
		gnRegistrationTop As Integer = 100
	#tag EndProperty

	#tag Property, Flags = &h0
		gnSearchHeight As Integer = 690
	#tag EndProperty

	#tag Property, Flags = &h0
		gnSearchLeft As Integer = 100
	#tag EndProperty

	#tag Property, Flags = &h0
		gnSearchTop As Integer = 100
	#tag EndProperty

	#tag Property, Flags = &h0
		gnSearchWidth As Integer = 1075
	#tag EndProperty

	#tag Property, Flags = &h0
		gPrintSettings As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gRS As RecordSet
	#tag EndProperty

	#tag Property, Flags = &h0
		grsHold As RecordSet
	#tag EndProperty

	#tag Property, Flags = &h0
		gsDefaultEvent As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsDefaultEventName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsFromEmail As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsFromName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsOS As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsPathDelimiter As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsReportName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsSMTPPassword As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsSMTPServer As String = "aspe.zmailcloud.com"
	#tag EndProperty

	#tag Property, Flags = &h0
		gsSMTPUserID As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsSql As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gsTempName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		gTempPID As Integer
	#tag EndProperty


	#tag Constant, Name = cnBulkEmailEmailPort, Type = Double, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"465"
	#tag EndConstant

	#tag Constant, Name = csBulkEmailSMTPPassword, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"wolfmail"
	#tag EndConstant

	#tag Constant, Name = csBulkMailSMTPServer, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"relay.jangosmtp.net"
	#tag EndConstant

	#tag Constant, Name = csBulkMailSMTPUserID, Type = String, Dynamic = False, Default = \"", Scope = Public
		#Tag Instance, Platform = Any, Language = Default, Definition  = \"aspechamp"
	#tag EndConstant

	#tag Constant, Name = Untitled, Type = , Dynamic = False, Default = \"", Scope = Public
	#tag EndConstant

	#tag Constant, Name = Untitled1, Type = , Dynamic = False, Default = \"", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="gbClose"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gbGroupSent"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gbLookUpPID"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gbOrderIsARegistration"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gbPostPayment"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gbRegistrationActive"
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gCurrentUser"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gdPayment"
			Group="Behavior"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gdTotalDue"
			Group="Behavior"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnEmailPort"
			Group="Behavior"
			InitialValue="465"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnFinanceHeight"
			Group="Behavior"
			InitialValue="600"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnFinanceLeft"
			Group="Behavior"
			InitialValue="20"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnFInanceTop"
			Group="Behavior"
			InitialValue="50"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnFinanceWidth"
			Group="Behavior"
			InitialValue="980"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnHomeLeft"
			Group="Behavior"
			InitialValue="10"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnHomeTop"
			Group="Behavior"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnOrderID"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnOrderLeft"
			Group="Behavior"
			InitialValue="200"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnOrderTop"
			Group="Behavior"
			InitialValue="200"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnPeopleLeft"
			Group="Behavior"
			InitialValue="200"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnPeopleTop"
			Group="Behavior"
			InitialValue="200"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnRegistrationLeft"
			Group="Behavior"
			InitialValue="10"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnRegistrationTop"
			Group="Behavior"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnSearchHeight"
			Group="Behavior"
			InitialValue="690"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnSearchLeft"
			Group="Behavior"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnSearchTop"
			Group="Behavior"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gnSearchWidth"
			Group="Behavior"
			InitialValue="1075"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gPrintSettings"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsDefaultEvent"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsDefaultEventName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsFromEmail"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsFromName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsOS"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsPathDelimiter"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsReportName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsSMTPPassword"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsSMTPServer"
			Group="Behavior"
			InitialValue="aspe.zmailcloud.com"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsSMTPUserID"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsSql"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gsTempName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="gTempPID"
			Group="Behavior"
			Type="Integer"
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
