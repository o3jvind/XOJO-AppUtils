#tag Module
Protected Module PreferencesTest
	#tag Method, Flags = &h0
		Sub RunAllPreferencesTests()
		  If App.prefs.Connect() Then
		    System.DebugLog("Connected successfully")
		    
		    // Test String
		    App.prefs.SetString("username", "JohnDoe")
		    System.DebugLog("String Test: " + App.prefs.GetString("username", "Guest"))
		    
		    // Test Boolean
		    App.prefs.SetBoolean("notifications_enabled", True)
		    System.DebugLog("Boolean Test: " + App.prefs.GetBoolean("notifications_enabled", False).ToString)
		    
		    // Test Integer32
		    App.prefs.SetInteger32("user_id", 12345)
		    System.DebugLog("Integer32 Test: " + App.prefs.GetInteger32("user_id", 0).ToString)
		    
		    // Test Integer64
		    App.prefs.SetInteger64("large_number", 9223372036854775807)
		    System.DebugLog("Integer64 Test: " + App.prefs.GetInteger64("large_number", 0).ToString)
		    
		    // Test Double
		    App.prefs.SetDouble("price", 99.99)
		    System.DebugLog("Double Test: " + App.prefs.GetDouble("price", 0.0).ToString)
		    
		    // Test Single
		    App.prefs.SetSingle("temperature", 36.6)
		    System.DebugLog("Single Test: " + Format(App.prefs.GetSingle("temperature", 0.0), "0.0"))
		    
		    // Test Color
		    Var testColor As Color = &cFF0000 // Red
		    App.prefs.SetColor("theme_color", testColor)
		    Var retrievedColor As Color = App.prefs.GetColor("theme_color", &c000000)
		    System.DebugLog("Color Test: " + retrievedColor.Red.ToString + "," + retrievedColor.Green.ToString + "," + retrievedColor.Blue.ToString)
		    
		    // Test DateTime
		    Var now As DateTime = DateTime.Now
		    App.prefs.SetDateTime("last_login", now)
		    Var retrievedDate As DateTime = App.prefs.GetDateTime("last_login", Nil)
		    System.DebugLog("DateTime Test: " + retrievedDate.SQLDateTime)
		    
		    // Test Array
		    Var testArray() As Variant
		    testArray.Add("Item 1")
		    testArray.Add("Item 2")
		    testArray.Add(123)
		    App.prefs.SetArray("recent_items", testArray)
		    Var retrievedArray As Variant = App.prefs.GetArray("recent_items", Nil)
		    If retrievedArray <> Nil Then
		      Var arr() As Variant = retrievedArray
		      System.DebugLog("Array Test: Count = " + arr.Count.ToString)
		      For i As Integer = 0 To arr.LastIndex
		        System.DebugLog("  Array Item " + i.ToString + ": " + arr(i).StringValue)
		      Next
		    End If
		    
		    // Test Dictionary
		    Var testDict As New Dictionary
		    testDict.Value("name") = "John"
		    testDict.Value("age") = 30
		    testDict.Value("city") = "New York"
		    App.prefs.SetDictionary("user_info", testDict)
		    Var retrievedDict As Dictionary = App.prefs.GetDictionary("user_info", Nil)
		    If retrievedDict <> Nil Then
		      System.DebugLog("Dictionary Test: Count = " + retrievedDict.Count.ToString)
		      For Each key As Variant In retrievedDict.Keys
		        System.DebugLog("  Dictionary Item " + key.StringValue + ": " + retrievedDict.Value(key).StringValue)
		      Next
		    End If
		    
		    // Test Picture (requires a picture to test)
		    TestPreferencesPicture()
		    
		    // Test DeleteKey
		    TestPreferencesDeleteKey()
		    
		    System.DebugLog("All tests completed")
		  Else
		    System.DebugLog("Failed to connect to database")
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TestPreferencesDeleteKey()
		  // Create a test entry
		  App.prefs.SetString("test_delete", "This will be deleted")
		  System.DebugLog("Created entry: " + App.prefs.GetString("test_delete", "Not found"))
		  
		  // Delete the entry
		  App.prefs.DeleteKey("test_delete")
		  System.DebugLog("After deletion: " + App.prefs.GetString("test_delete", "Not found"))
		  
		  // Verify the entry is gone
		  If App.prefs.KeyExists("test_delete") Then
		    System.DebugLog("DeleteKey Test: Failed - Key still exists")
		  Else
		    System.DebugLog("DeleteKey Test: Passed - Key was deleted successfully")
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub TestPreferencesPicture()
		  // Create a test picture
		  Var testPic As New Picture(100, 100)
		  Var g As Graphics = testPic.Graphics
		  
		  // Draw something on the picture
		  g.DrawingColor = &cFF0000 // Red
		  g.FillRectangle(0, 0, 100, 100)
		  g.DrawingColor = &c0000FF // Blue
		  g.FillOval(25, 25, 50, 50)
		  
		  // Store the picture
		  App.prefs.SetPicture("test_picture", testPic)
		  System.DebugLog("Picture stored successfully")
		  
		  // Retrieve the picture
		  Var retrievedPic As Picture = App.prefs.GetPicture("test_picture", Nil)
		  If retrievedPic <> Nil Then
		    System.DebugLog("Picture Test: Width = " + retrievedPic.Width.ToString + ", Height = " + retrievedPic.Height.ToString)
		  Else
		    System.DebugLog("Picture Test: Failed to retrieve picture")
		  End If
		End Sub
	#tag EndMethod


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
