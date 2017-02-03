#tag Class
Protected Class App
Inherits WebApplication
	#tag Event
		Sub Open(args() as String)
		  self.Security.FrameEmbedding = WebAppSecurityOptions.FrameOptions.Allow
		  App.AutoQuit = True
		  App.Timeout = 1
		  App.SessionTimeout = 1
		  
		End Sub
	#tag EndEvent

	#tag Event
		Function UnhandledException(Error As RuntimeException) As Boolean
		  Try
		    dim log as string
		    dim d as new date
		    log = log + d.SQLDateTime + EndOfLine
		    log = log + Error.Message + EndOfLine + error.Reason + join(error.stack, EndOfLine)
		    
		    'f= GetFolderItem("log.txt")
		    Dim f As FolderItem = GetFolderItem("AppErrorLog.txt")
		    Dim t as TextOutputStream
		    If f <> Nil then
		      t = TextOutputStream.Append(f)
		      t.WriteLine(log)
		      t.Close
		    End If 
		    Return true
		  Catch
		    Return True
		  end
		  
		  
		  
		  'dim ls as String
		  '
		  'ls = "Runtime Exception: " + Error.Type + EndOfLine + _
		  '"                           Reason: " + error.Reason + EndOfLine + _
		  '"                     Error Number: " + error.ErrorNumber.ToText + EndOfLine + _
		  '"                          Message: " + error.Message + EndOfLine + _
		  '"                            Stack: " + Join(error.Stack())
		  'WriteLog(ls)
		  '
		  'Return False
		End Function
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub WriteLog(lsText as String)
		  dim f as folderitem
		  
		  f= GetFolderItem("log.txt") 'GetFolderItem("").parent.child("log.txt")
		  
		  dim tos as TextOutputStream
		  Dim ld as New Date
		  
		  if f.exists then
		    tos = TextOutputStream.Append(f)
		  else
		    tos = TextOutputStream.create(f)
		  end
		  
		  tos.WriteLine ld.ShortDate + " " + ld.LongTime  + " - " + lsText
		  tos.close
		  
		End Sub
	#tag EndMethod


	#tag ViewBehavior
	#tag EndViewBehavior
End Class
#tag EndClass
