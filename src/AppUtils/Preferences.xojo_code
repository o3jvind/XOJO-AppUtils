#tag Class
Protected Class Preferences
	#tag Method, Flags = &h0
		Sub ClearAll()
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    mDatabase.ExecuteSQL("DELETE FROM preferences")
		    
		    ' Clear the memory cache
		    mPrefsCache = New Dictionary
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.ClearAll: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Connect() As Boolean
		  Try
		    ' Check if we're already connected
		    If mConnected And mDatabase <> Nil Then
		      Return True
		    End If
		    
		    ' Create a new database instance
		    mDatabase = New SQLiteDatabase
		    
		    ' Set the database file
		    mDatabase.DatabaseFile = mDatabaseFile
		    
		    ' Debug the database file path
		    System.DebugLog("Database file path: " + mDatabaseFile.NativePath)
		    
		    ' Make sure the parent folder exists
		    If Not mDatabaseFile.Parent.Exists Then
		      mDatabaseFile.Parent.CreateAsFolder
		      System.DebugLog("Created parent folder: " + mDatabaseFile.Parent.NativePath)
		    End If
		    
		    ' Check if the database file exists, create it if it doesn't
		    If Not mDatabaseFile.Exists Then
		      System.DebugLog("Database file doesn't exist, creating it...")
		      
		      ' Use the built-in method to create a proper SQLite database file
		      If mDatabase.CreateDatabaseFile Then
		        System.DebugLog("Database file created successfully")
		      Else
		        System.DebugLog("Failed to create database file: " + mDatabase.ErrorMessage)
		        Return False
		      End If
		    End If
		    
		    ' Connect to the database
		    If mDatabase.Connect Then
		      System.DebugLog("Successfully connected to database")
		      
		      ' Create tables if needed
		      If CreateTablesIfNeeded Then
		        mConnected = True
		        System.DebugLog("Tables created successfully")
		        
		        ' Load all preferences into memory cache
		        LoadPreferencesIntoCache()
		        
		        Return True
		      Else
		        System.DebugLog("Failed to create tables: " + mDatabase.ErrorMessage)
		      End If
		    Else
		      System.DebugLog("Failed to connect to database: " + mDatabase.ErrorMessage)
		    End If
		    
		    ' If we get here, connection failed
		    mDatabase = Nil
		    mConnected = False
		    Return False
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.Connect exception: " + e.Message)
		    mDatabase = Nil
		    mConnected = False
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(appName As String = "")
		  mConnected = False
		  mDatabase = Nil
		  
		  ' Initialize the memory cache
		  mPrefsCache = New Dictionary
		  
		  ' Set default app name if not provided
		  If appName = "" Then
		    mAppName = "com.app.preferences"
		  Else
		    mAppName = appName
		  End If
		  
		  ' Initialize database file path
		  If Not InitializeDatabase() Then
		    System.DebugLog("Preferences.Constructor: Failed to initialize database")
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CreateTablesIfNeeded() As Boolean
		  Try
		    If mDatabase <> Nil Then
		      ' Create the preferences table if it doesn't exist
		      mDatabase.ExecuteSQL("CREATE TABLE IF NOT EXISTS preferences (" + _
		      "key TEXT PRIMARY KEY, " + _
		      "value BLOB, " + _
		      "type TEXT)")
		      Return True
		    End If
		    
		    Return False
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.CreateTablesIfNeeded: " + e.Message)
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeleteKey(key As String)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    mDatabase.ExecuteSQL("DELETE FROM preferences WHERE key = ?", key)
		    
		    ' Remove from memory cache
		    If mPrefsCache.HasKey(key) Then
		      mPrefsCache.Remove(key)
		    End If
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.DeleteKey: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Disconnect()
		  Try
		    If mConnected And mDatabase <> Nil Then
		      mDatabase.Close
		    End If
		  Catch e As RuntimeException
		    ' Ignore errors on disconnect
		  End Try
		  
		  mDatabase = Nil
		  mConnected = False
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetArray(key As String, defaultValue As Variant = Nil) As Variant
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "array" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "array")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      
		      ' Parse JSON string to array
		      Var json As New JSONItem(strValue)
		      
		      ' Create an array to hold the values
		      Var tempArray() As Variant
		      
		      ' Get the number of items in the JSON array
		      Var count As Integer = json.Count
		      
		      ' Copy values from JSON to array
		      For i As Integer = 0 To count - 1
		        tempArray.Add(json.Value(i.ToString))
		      Next
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "array"
		      cacheItem.Value("value") = tempArray
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return tempArray
		    End If
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.GetArray: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetBoolean(key As String, defaultValue As Boolean = False) As Boolean
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "boolean" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "boolean")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      Var result As Boolean = (strValue = "1")
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "boolean"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetBoolean: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetColor(key As String, defaultValue As Color = &c000000) As Color
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "color" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "color")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      
		      ' Parse hex color string
		      If strValue.Left(2) = "&c" And strValue.Length = 8 Then
		        Var hexColor As String = strValue.Mid(3)
		        Var r As Integer = Integer.FromHex(hexColor.Left(2))
		        Var g As Integer = Integer.FromHex(hexColor.Mid(3, 2))
		        Var b As Integer = Integer.FromHex(hexColor.Right(2))
		        
		        Var result As Color = RGB(r, g, b)
		        
		        ' Store in cache
		        Var cacheItem As New Dictionary
		        cacheItem.Value("type") = "color"
		        cacheItem.Value("value") = result
		        mPrefsCache.Value(key) = cacheItem
		        
		        Return result
		      End If
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetColor: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDatabaseFilePath() As String
		  If mDatabaseFile <> Nil Then
		    Return mDatabaseFile.NativePath
		  End If
		  
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDateTime(key As String, defaultValue As DateTime = Nil) As DateTime
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "datetime" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "datetime")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      
		      ' Parse ISO 8601 string to DateTime
		      Try
		        Var dt As DateTime = DateTime.FromString(strValue)
		        
		        ' Store in cache
		        Var cacheItem As New Dictionary
		        cacheItem.Value("type") = "datetime"
		        cacheItem.Value("value") = dt
		        mPrefsCache.Value(key) = cacheItem
		        
		        Return dt
		      Catch
		        ' If parsing fails, return default
		      End Try
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetDateTime: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDictionary(key As String, defaultValue As Dictionary = Nil) As Dictionary
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "dictionary" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "dictionary")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      
		      ' Parse JSON string to dictionary
		      Var json As New JSONItem(strValue)
		      Var result As New Dictionary
		      
		      ' Convert JSON to Dictionary
		      For Each k As String In json.Names
		        result.Value(k) = json.Value(k)
		      Next
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "dictionary"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.GetDictionary: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetDouble(key As String, defaultValue As Double = 0.0) As Double
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "double" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "double")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      Var result As Double = Double.FromString(strValue)
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "double"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetDouble: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetInteger32(key As String, defaultValue As Integer = 0) As Integer
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "integer32" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "integer32")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      Var result As Integer = Integer.FromString(strValue)
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "integer32"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetInteger32: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetInteger64(key As String, defaultValue As Int64 = 0) As Int64
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "integer64" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "integer64")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      Var result As Int64 = Int64.FromString(strValue)
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "integer64"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetInteger64: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetPicture(key As String, defaultValue As Picture = Nil) As Picture
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "picture" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "picture")
		    If Not rs.AfterLastRow Then
		      Var data As String = rs.Column("value").StringValue
		      
		      If data <> "" Then
		        ' Convert data to Picture
		        Var result As Picture = Picture.FromData(data)
		        
		        If result <> Nil Then
		          ' Store in cache
		          Var cacheItem As New Dictionary
		          cacheItem.Value("type") = "picture"
		          cacheItem.Value("value") = result
		          mPrefsCache.Value(key) = cacheItem
		          
		          Return result
		        End If
		      End If
		    End If
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.GetPicture: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetPlatformSpecificPath() As FolderItem
		  Var prefsFolder As FolderItem
		  Var dbFile As FolderItem
		  
		  Try
		    #If TargetMacOS Then
		      ' macOS: ~/Library/Application Support/com.your.app/
		      prefsFolder = SpecialFolder.ApplicationData.Child(mAppName)
		    #ElseIf TargetWindows Then
		      ' Windows: %APPDATA%\com.your.app\
		      prefsFolder = SpecialFolder.ApplicationData.Child(mAppName)
		    #ElseIf TargetLinux Then
		      ' Linux: ~/.config/appname/
		      prefsFolder = SpecialFolder.ApplicationData.Child(".config").Child(mAppName)
		    #EndIf
		    
		    ' Create the directory if it doesn't exist
		    If Not prefsFolder.Exists Then
		      prefsFolder.CreateAsFolder
		    End If
		    
		    ' Create the database file path
		    dbFile = prefsFolder.Child("preferences.sqlite")
		    
		    Return dbFile
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.GetPlatformSpecificPath: " + e.Message)
		    Return Nil
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetSingle(key As String, defaultValue As Single = 0.0) As Single
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "single" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "single")
		    If Not rs.AfterLastRow Then
		      Var strValue As String = rs.Column("value").StringValue
		      Var result As Single = Single.FromString(strValue)
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "single"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetSingle: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetString(key As String, defaultValue As String = "") As String
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Var cacheItem As Dictionary = mPrefsCache.Value(key)
		    If cacheItem.Value("type") = "string" Then
		      Return cacheItem.Value("value")
		    End If
		  End If
		  
		  ' If not in cache, try to get from database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return defaultValue
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT value FROM preferences WHERE key = ? AND type = ?", key, "string")
		    If Not rs.AfterLastRow Then
		      Var result As String = rs.Column("value").StringValue
		      
		      ' Store in cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "string"
		      cacheItem.Value("value") = result
		      mPrefsCache.Value(key) = cacheItem
		      
		      Return result
		    End If
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.GetString: " + e.Message)
		  End Try
		  
		  Return defaultValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function InitializeDatabase() As Boolean
		  Try
		    ' Get platform-specific path
		    mDatabaseFile = GetPlatformSpecificPath()
		    
		    ' Check if we got a valid folder item
		    If mDatabaseFile <> Nil Then
		      Return True
		    End If
		    
		    Return False
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.InitializeDatabase: " + e.Message)
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsConnected() As Boolean
		  Return mConnected And mDatabase <> Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function KeyExists(key As String) As Boolean
		  ' Check memory cache first
		  If mPrefsCache.HasKey(key) Then
		    Return True
		  End If
		  
		  ' If not in cache, check database
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return False
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT key FROM preferences WHERE key = ?", key)
		    Return Not rs.AfterLastRow
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.KeyExists: " + e.Message)
		    Return False
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub LoadPreferencesIntoCache()
		  ' Clear existing cache
		  mPrefsCache = New Dictionary
		  
		  Try
		    ' Get all preferences from database
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT key, value, type FROM preferences")
		    
		    ' Loop through all rows and add to cache
		    While Not rs.AfterLastRow
		      Var key As String = rs.Column("key").StringValue
		      Var type As String = rs.Column("type").StringValue
		      
		      ' Create cache item
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = type
		      
		      ' Convert value based on type
		      Select Case type
		      Case "string"
		        cacheItem.Value("value") = rs.Column("value").StringValue
		        
		      Case "boolean"
		        cacheItem.Value("value") = (rs.Column("value").StringValue = "1")
		        
		      Case "integer32"
		        cacheItem.Value("value") = Integer.FromString(rs.Column("value").StringValue)
		        
		      Case "integer64"
		        cacheItem.Value("value") = Int64.FromString(rs.Column("value").StringValue)
		        
		      Case "double"
		        cacheItem.Value("value") = Double.FromString(rs.Column("value").StringValue)
		        
		      Case "single"
		        cacheItem.Value("value") = Single.FromString(rs.Column("value").StringValue)
		        
		      Case "color"
		        Var value As String = rs.Column("value").StringValue
		        If value.Left(2) = "&c" And value.Length = 8 Then
		          Var hexColor As String = value.Mid(3)
		          Var r As Integer = Integer.FromHex(hexColor.Left(2))
		          Var g As Integer = Integer.FromHex(hexColor.Mid(3, 2))
		          Var b As Integer = Integer.FromHex(hexColor.Right(2))
		          cacheItem.Value("value") = RGB(r, g, b)
		        End If
		        
		      Case "datetime"
		        Try
		          Var dtStr As String = rs.Column("value").StringValue
		          cacheItem.Value("value") = DateTime.FromString(dtStr)
		        Catch
		          ' Skip if datetime parsing fails
		          rs.MoveToNextRow
		          Continue
		        End Try
		        
		      Case "array"
		        Try
		          Var json As New JSONItem(rs.Column("value").StringValue)
		          
		          ' Create a new array to hold the values
		          Var tempArray() As Variant
		          
		          ' Get the number of items in the JSON array
		          Var count As Integer = json.Count
		          
		          ' Copy values from JSON to array
		          For i As Integer = 0 To count - 1
		            tempArray.Add(json.Value(i.ToString))
		          Next
		          
		          cacheItem.Value("value") = tempArray
		        Catch
		          ' Skip if JSON parsing fails
		          rs.MoveToNextRow
		          Continue
		        End Try
		        
		      Case "dictionary"
		        Try
		          Var json As New JSONItem(rs.Column("value").StringValue)
		          Var dict As New Dictionary
		          
		          For Each k As String In json.Names
		            dict.Value(k) = json.Value(k)
		          Next
		          
		          cacheItem.Value("value") = dict
		        Catch
		          ' Skip if JSON parsing fails
		          rs.MoveToNextRow
		          Continue
		        End Try
		        
		      Case "picture"
		        Try
		          Var data As String = rs.Column("value").StringValue
		          If data <> "" Then
		            Var pic As Picture = Picture.FromData(data)
		            If pic <> Nil Then
		              cacheItem.Value("value") = pic
		            Else
		              ' Skip if picture creation fails
		              rs.MoveToNextRow
		              Continue
		            End If
		          Else
		            ' Skip if data is empty
		            rs.MoveToNextRow
		            Continue
		          End If
		        Catch
		          ' Skip if picture parsing fails
		          rs.MoveToNextRow
		          Continue
		        End Try
		        
		      End Select
		      
		      ' Add to cache
		      mPrefsCache.Value(key) = cacheItem
		      
		      rs.MoveToNextRow
		    Wend
		    
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.LoadPreferencesIntoCache: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function PreferencesCount() As Integer
		  If Not IsConnected Then
		    If Not Connect Then
		      System.DebugLog("Preferences.PreferencesCount: Could not connect to database.")
		      Return -1
		    End If
		  End If
		  
		  Try
		    Var rs As RowSet = mDatabase.SelectSQL("SELECT COUNT(*) AS total FROM preferences")
		    Return rs.Column("total").IntegerValue
		  Catch e As DatabaseException
		    System.DebugLog("Preferences.PreferencesCount: " + e.Message)
		    Return -1
		  End Try
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetArray(key As String, value As Variant)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    ' Convert array to JSON string
		    Var json As New JSONItem
		    
		    ' Check if value is an array using VarType
		    If (VarType(value) And Variant.TypeArray) = Variant.TypeArray Then
		      ' Get array elements
		      Var arr() As Variant = value
		      
		      ' Add array elements to JSON
		      For i As Integer = 0 To arr.LastIndex
		        json.Value(i.ToString) = arr(i)
		      Next
		      
		      Var strValue As String = json.ToString
		      
		      ' Update database
		      If KeyExists(key) Then
		        mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "array", key)
		      Else
		        mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "array")
		      End If
		      
		      ' Update memory cache
		      Var cacheItem As New Dictionary
		      cacheItem.Value("type") = "array"
		      cacheItem.Value("value") = arr
		      mPrefsCache.Value(key) = cacheItem
		    End If
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.SetArray: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetBoolean(key As String, value As Boolean)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    Var strValue As String = If(value, "1", "0")
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "boolean", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "boolean")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "boolean"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetBoolean: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetColor(key As String, value As Color)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    ' Convert color to hex string
		    Var strValue As String = "&c" + value.Red.ToHex(2) + value.Green.ToHex(2) + value.Blue.ToHex(2)
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "color", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "color")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "color"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetColor: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetDateTime(key As String, value As DateTime)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    ' Convert DateTime to ISO 8601 string
		    Var strValue As String = value.SQLDateTime
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "datetime", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "datetime")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "datetime"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetDateTime: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetDictionary(key As String, value As Dictionary)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    ' Convert dictionary to JSON string
		    Var json As New JSONItem
		    
		    ' Add dictionary entries to JSON
		    For Each k As Variant In value.Keys
		      json.Value(k.StringValue) = value.Value(k)
		    Next
		    
		    Var strValue As String = json.ToString
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "dictionary", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "dictionary")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "dictionary"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.SetDictionary: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetDouble(key As String, value As Double)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    Var strValue As String = value.ToString
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "double", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "double")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "double"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetDouble: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetInteger32(key As String, value As Integer)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    Var strValue As String = value.ToString
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "integer32", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "integer32")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "integer32"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetInteger32: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetInteger64(key As String, value As Int64)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    Var strValue As String = value.ToString
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "integer64", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "integer64")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "integer64"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetInteger64: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetPicture(key As String, value As Picture)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    ' Convert picture to string data
		    Var data As String = value.ToData(Picture.Formats.PNG)
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", data, "picture", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, data, "picture")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "picture"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As RuntimeException
		    System.DebugLog("Preferences.SetPicture: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetSingle(key As String, value As Single)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    Var strValue As String = value.ToString
		    
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", strValue, "single", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, strValue, "single")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "single"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetSingle: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SetString(key As String, value As String)
		  If Not IsConnected() Then
		    If Not Connect() Then
		      LastError = "Could not connect to database."
		      Raise New RuntimeException(LastError)
		      Return
		    End If
		  End If
		  
		  Try
		    ' Update database
		    If KeyExists(key) Then
		      mDatabase.ExecuteSQL("UPDATE preferences SET value = ?, type = ? WHERE key = ?", value, "string", key)
		    Else
		      mDatabase.ExecuteSQL("INSERT INTO preferences (key, value, type) VALUES (?, ?, ?)", key, value, "string")
		    End If
		    
		    ' Update memory cache
		    Var cacheItem As New Dictionary
		    cacheItem.Value("type") = "string"
		    cacheItem.Value("value") = value
		    mPrefsCache.Value(key) = cacheItem
		    
		  Catch e As DatabaseException
		    LastError = e.Message
		    System.DebugLog("Preferences.SetString: " + e.Message)
		  End Try
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		LastError As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mAppName As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mConnected As Boolean
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDatabase As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDatabaseFile As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mPrefsCache As Dictionary
	#tag EndProperty


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
		#tag ViewProperty
			Name="LastError"
			Visible=false
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
