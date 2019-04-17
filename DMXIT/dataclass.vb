Imports MongoDB.Bson
Imports MongoDB.Driver
Imports System.Configuration
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class dataclass
    Public db As IMongoDatabase
    Public coll As IMongoCollection(Of BsonDocument)
    Public coll2 As IMongoCollection(Of BsonDocument)
    Public coll3 As IMongoCollection(Of BsonDocument)
    Public coll4 As IMongoCollection(Of BsonDocument)
    Public coll5 As IMongoCollection(Of BsonDocument)

    Public Function InitiateDb()
        Dim conn As String = ConfigurationManager.AppSettings("mongoDb").ToString()
        Dim client As New MongoClient(conn)
        db = client.GetDatabase("myDmxLibrary")

        coll = db.GetCollection(Of BsonDocument)("fixtures")
        coll2 = db.GetCollection(Of BsonDocument)("scenes")
        coll3 = db.GetCollection(Of BsonDocument)("layouts")
        coll4 = db.GetCollection(Of BsonDocument)("slices")
        coll5 = db.GetCollection(Of BsonDocument)("slicecalls")

        Return 0
    End Function


    Public Function LoadComboDefaultRanges(max As Integer) As DataTable
        Dim t As New DataTable

        t.Columns.Add("item")

        For x = 0 To max
            t.Rows.Add(x)
        Next

        LoadComboDefaultRanges = t
    End Function

    Public Function LoadSliceCombos(obj As String) As DataTable
        Dim t As New DataTable

        ' values for duration combo
        Dim s1 As String = "0,.5,1,2,3,4,5,6,7,8,9,10,20,30,60"

        ' values for fade combo
        Dim s2 As String = "yes, no"

        t.Columns.Add("alias")
        t.Columns.Add("item")

        Dim a1 As String() = Split(s1, ",")
        Dim a2 As String() = Split(s2, ",")

        Select Case obj
            Case "duration"
                For Each v As String In a1
                    Dim r As DataRow = t.NewRow()
                    r("item") = v
                    r("alias") = v & " seconds"
                    t.Rows.Add(r)
                Next

            Case "fade"
                For Each v As String In a2
                    Dim r As DataRow = t.NewRow()
                    r("item") = v
                    r("alias") = v
                    t.Rows.Add(r)
                Next

        End Select

        Return t
    End Function


    Public Async Function DoesRecordExist(collectionname As String, fieldname As String, fieldvalue As String) As Task(Of String)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim test As Integer
        Dim id As New ObjectId

        vCol = db.GetCollection(Of BsonDocument)(collectionname)

        Dim query As BsonDocument
        query = New BsonDocument(fieldname, fieldvalue)

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()
        If myList.Count = 0 Then
            test = 0
        Else
            test = 1
        End If

        For Each Doc As BsonDocument In myList
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value
                If ColName = "_id" Then
                    id = ColValue
                End If
            Next
        Next

        Return id.ToString

    End Function

    Public Async Function GetDevices() As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        Dim query As BsonDocument
        query = New BsonDocument()

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("_id")
        Table.Columns.Add("name")
        Table.Columns.Add("manufacturer")
        Table.Columns.Add("model")
        Table.Columns.Add("channels")
        Table.Columns.Add("quantity")


        Dim cnt As Integer = -2
        For Each Doc As BsonDocument In myList
            cnt += 1
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value
                If Not ColName = "configuration" Then ' remove default column
                    Dr(ColName) = ColValue
                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function GetScenes() As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("scenes")


        Dim query As BsonDocument
        query = New BsonDocument()

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("name")
        Table.Columns.Add("type")
        Table.Columns.Add("actions")


        Dim cnt As Integer = -2
        For Each Doc As BsonDocument In myList
            cnt += 1
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value
                If Not ColName = "_id" Then ' remove default column
                    If Dr.Table.Columns.Contains(ColName) = True Then
                        Dr(ColName) = ColValue
                    Else
                        Table.Columns.Add(ColName)
                        Dr(ColName) = ColValue
                    End If

                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function GetLayouts() As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("layouts")


        Dim query As BsonDocument
        query = New BsonDocument()

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("id")
        Table.Columns.Add("name")
        Table.Columns.Add("location")
        Table.Columns.Add("devicegroups")
        Table.Columns.Add("dmxid")


        Dim cnt As Integer = -2
        For Each Doc As BsonDocument In myList
            cnt += 1
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value

                If Dr.Table.Columns.Contains(ColName) = True Then
                    Dr(ColName) = ColValue
                Else
                    Table.Columns.Add(ColName)
                    Dr(ColName) = ColValue
                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function LoadFixures() As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        Dim query As BsonDocument
        query = New BsonDocument()

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("name")

        For Each Doc As BsonDocument In myList
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value
                If ColName = "name" Then
                    Dr(ColName) = ColValue
                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function GetDeviceConfiguration(keyname As String, keyvalue As String) As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim o As JObject
        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        Dim query As BsonDocument

        Select Case keyname
            Case "_id"
                query = New BsonDocument("_id", New ObjectId(keyvalue))
            Case Else
                query = New BsonDocument(keyname, keyvalue)
        End Select


        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()


        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("startchannel")
        Table.Columns.Add("dmxvalue")
        Table.Columns.Add("dmxfunctioncategory")
        Table.Columns.Add("dmxfunction")


        Dim cnt As Integer = -2
        For Each Doc As BsonDocument In myList
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value
                If ColName = "configuration" Then
                    Dim sArray = ColValue.ToString.ToCharArray()
                    Dim device As List(Of DeviceConfig) = JsonConvert.DeserializeObject(Of List(Of DeviceConfig))(ColValue.ToString)
                    For Each item As DeviceConfig In device
                        Dim Dr As DataRow = Table.NewRow()
                        Dr("startchannel") = item.startchannel
                        Dr("dmxvalue") = item.dmxvalue
                        Dr("dmxfunctioncategory") = item.dmxfunctioncategory
                        Dr("dmxfunction") = item.dmxfunction
                        Table.Rows.Add(Dr)
                    Next
                End If
            Next

        Next

        Return Table

    End Function

    Public Async Function GetLayoutConfiguration(keyname As String, keyvalue As String) As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim o As JObject
        vCol = db.GetCollection(Of BsonDocument)("layouts")


        Dim query As BsonDocument

        Select Case keyname
            Case "_id"
                query = New BsonDocument("_id", New ObjectId(keyvalue))
            Case Else
                query = New BsonDocument(keyname, keyvalue)
        End Select


        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()


        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("devicename")
        Table.Columns.Add("devicealias")
        Table.Columns.Add("dmxid")
        Table.Columns.Add("devicegroups")
        Table.Columns.Add("deviceref") ' used for easy reference in combo selection

        Dim cnt As Integer = -2
        For Each Doc As BsonDocument In myList
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value
                If ColName = "devices" Then
                    Dim sArray = ColValue.ToString.ToCharArray()
                    Dim device As List(Of LayoutConfig) = JsonConvert.DeserializeObject(Of List(Of LayoutConfig))(ColValue.ToString)
                    For Each item As LayoutConfig In device
                        Dim Dr As DataRow = Table.NewRow()
                        Dr("devicename") = item.devicename
                        Dr("devicealias") = item.devicealias
                        Dr("dmxid") = item.dmxid
                        Dr("devicegroups") = item.devicegroups
                        Dr("deviceref") = item.devicename & ": " & item.devicealias & " [" & "dmx " & item.dmxid & "]"
                        Table.Rows.Add(Dr)
                    Next
                End If
            Next

        Next

        Return Table

    End Function

    Public Async Function GetSlice(keyname As String, keyvalue As String) As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim o As JObject
        vCol = db.GetCollection(Of BsonDocument)("slices")


        Dim query As BsonDocument

        Select Case keyname
            Case "_id"
                query = New BsonDocument("_id", New ObjectId(keyvalue))
            Case Else
                query = New BsonDocument(keyname, keyvalue)
        End Select


        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()


        ' get all records
        Dim Table As New DataTable
        Table.Columns.Add("_id")
        Table.Columns.Add("name")
        Table.Columns.Add("duration")
        Table.Columns.Add("fade")
        Table.Columns.Add("nextslice")
        Table.Columns.Add("dmxstring")


        For Each Doc As BsonDocument In myList
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value

                If Table.Columns.Contains(ColName) = True Then
                    Dr(ColName) = ColValue
                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function GetSlices() As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("slices")


        Dim query As BsonDocument
        query = New BsonDocument()
        'query = New BsonDocument("layout", slayout)

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("_id")
        Table.Columns.Add("name")
        Table.Columns.Add("duration")
        Table.Columns.Add("fade")
        Table.Columns.Add("dmxstring")
        'Table.Columns.Add("layout")


        For Each Doc As BsonDocument In myList
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value

                If Dr.Table.Columns.Contains(ColName) = True Then
                    Dr(ColName) = ColValue
                Else
                    Table.Columns.Add(ColName)
                    Dr(ColName) = ColValue
                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function AddNewRecord(fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim r As BsonDocument = New BsonDocument
        Dim bsonlist As List(Of BsonDocument) = New List(Of BsonDocument)
        Dim temp As BsonDocument = New BsonDocument

        ' get collection
        vCol = db.GetCollection(Of BsonDocument)("fixtures")

        ' define nested configuration array
        Dim one As BsonDocument = New BsonDocument
        With one
            .Add("startchannel", 1)
            .Add("dmxvalue", 1)
            .Add("dmxfunctioncategory", "")
            .Add("dmxfunction", "")
        End With

        Dim test As New BsonArray
        With test
            .Add(one)
        End With

        ' create bsondocument
        r.Add("name", fieldvalue)
        r.Add("configuration", test)

        ' insert record
        vCol.InsertOne(r)

        Return 0

    End Function

    Public Async Function AddNewSceneRecord(fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim r As BsonDocument = New BsonDocument
        Dim bsonlist As List(Of BsonDocument) = New List(Of BsonDocument)
        Dim temp As BsonDocument = New BsonDocument

        ' get collection
        vCol = db.GetCollection(Of BsonDocument)("scenes")

        ' create bsondocument
        r.Add("name", fieldvalue)

        ' insert record
        vCol.InsertOne(r)

        Return 0

    End Function

    Public Async Function AddNewSliceCall(comment As String, position As String, contentid As Integer) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim b As BsonDocument = New BsonDocument
        Dim bsonlist As List(Of BsonDocument) = New List(Of BsonDocument)

        ' get collection
        vCol = db.GetCollection(Of BsonDocument)("slicecalls")

        ' create bsondocument
        b.Add("comment", comment)
        b.Add("position", position)
        b.Add("contentid", contentid)

        ' insert record
        vCol.InsertOne(b)

        Return 0

    End Function

    Public Async Function GetSliceCalls(contentid As Integer) As Task(Of DataTable)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("slicecalls")


        Dim query As BsonDocument
        query = New BsonDocument()
        query = New BsonDocument("contentid", contentid)

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' get all records

        Dim Table As New DataTable
        Table.Columns.Add("_id")
        Table.Columns.Add("comment")
        Table.Columns.Add("position")
        Table.Columns.Add("contentid")
        Table.Columns.Add("slicename")
        'Table.Columns.Add("layout")


        For Each Doc As BsonDocument In myList
            Dim Dr As DataRow = Table.NewRow()
            For Each element As BsonElement In Doc.Elements
                Dim ColName As String = element.Name
                Dim ColValue As BsonValue = element.Value

                If Dr.Table.Columns.Contains(ColName) = True Then
                    Dr(ColName) = ColValue
                End If
            Next
            Table.Rows.Add(Dr)
        Next

        Return Table

    End Function

    Public Async Function UpdateSliceCallRecord(keyvalue As String, dgindex As String, fieldname As String, fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("slicecalls")

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyvalue))

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$set", New BsonDocument(fieldname, fieldvalue)))

        Return 0

    End Function

    Public Async Function DeleteSliceCall(keyvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("slicecalls")


        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyvalue))

        ' update value if found
        vCol.DeleteMany(fltr)

        Return 0

    End Function

    Public Async Function AddNewLayoutRecord(fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)
        Dim r As BsonDocument = New BsonDocument
        Dim bsonlist As List(Of BsonDocument) = New List(Of BsonDocument)
        Dim temp As BsonDocument = New BsonDocument

        ' get collection
        vCol = db.GetCollection(Of BsonDocument)("layouts")

        ' define nested configuration array
        Dim one As BsonDocument = New BsonDocument
        With one
            .Add("devicename", "")
            .Add("devicealias", "")
            .Add("dmxid", 0)
            .Add("devicegroups", "")
        End With

        Dim test As New BsonArray
        With test
            .Add(one)
        End With

        ' create bsondocument
        r.Add("name", fieldvalue)
        r.Add("devices", test)

        ' insert record
        vCol.InsertOne(r)

        Return 0

    End Function

    Public Async Function AddNewConfigRecord(keyvalue As String) As Task(Of Integer)
        Dim r As BsonDocument = New BsonDocument
        Dim temp As BsonDocument = New BsonDocument
        Dim bsonlist As List(Of BsonDocument) = New List(Of BsonDocument)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        ' define nested configuration array
        Dim one As BsonDocument = New BsonDocument
        With one
            .Add("startchannel", 1)
            .Add("dmxvalue", 1)
            .Add("dmxfunctioncategory", "")
            .Add("dmxfunction", "")
        End With

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyvalue))
        vCol.UpdateOne(fltr, New BsonDocument("$push", New BsonDocument("configuration", one)))

        Return 0

    End Function

    Public Async Function saveNewSlice(slicename As String, duration As String, fade As String, nextslice As String, dmxdata() As Byte) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("slices")


        ' define nested configuration array
        Dim a As String = Text.Encoding.Default.GetString(dmxdata)

        Dim a1() As Byte = Text.Encoding.Default.GetBytes(a)

        Dim b As BsonDocument = New BsonDocument
        With b
            .Add("name", slicename)
            .Add("duration", duration)
            .Add("fade", fade)
            .Add("nextslice", nextslice)
            .Add("dmxstring", a)
        End With

        vCol.InsertOne(b)

        Return 0

    End Function

    Public Async Function updateSlice(sliceid As String, slicename As String, duration As String, fade As String, nextslice As String, dmxdata() As Byte) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("slices")


        ' encode dmx byte array to string
        Dim a As String = Text.Encoding.Default.GetString(dmxdata)

        ' define nested configuration array
        Dim b As BsonDocument = New BsonDocument
        With b
            .Add("name", slicename)
            .Add("duration", duration)
            .Add("fade", fade)
            .Add("nextslice", nextslice)
            .Add("dmxstring", a)
        End With

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(sliceid))

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$set", b))

        Return 0

    End Function

    Public Async Function AddDeviceToLayout(keyvalue As String, layoutid As String) As Task(Of Integer)
        Dim r As BsonDocument = New BsonDocument
        Dim temp As BsonDocument = New BsonDocument
        Dim bsonlist As List(Of BsonDocument) = New List(Of BsonDocument)
        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("layouts")


        ' define nested configuration array
        Dim one As BsonDocument = New BsonDocument
        With one
            .Add("devicename", keyvalue)
            .Add("devicealias", "")
            .Add("dmxid", 0)
            .Add("devicegroups", "")
        End With

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(layoutid))
        vCol.UpdateOne(fltr, New BsonDocument("$push", New BsonDocument("devices", one)))

        Return 0

    End Function

    Public Async Function UpdateRecord(keyname As String, fieldname As String, fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$set", New BsonDocument(fieldname, fieldvalue)))

        Return 0

    End Function

    Public Async Function UpdateScene(keyname As String, fieldvalue As Array) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("scenes")


        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))
        vCol.UpdateOne(fltr, New BsonDocument("$set", New BsonDocument("data", "test")))

        Return 0

    End Function

    Public Async Function UpdateConfigRecord(keyname As String, dgindex As Integer, fieldname As String, fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("fixtures")

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' define update document
        Dim doc1 As BsonDocument = New BsonDocument
        With doc1
            .Add("configuration." & dgindex & "." & fieldname, fieldvalue)
        End With

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$set", doc1))

        Return 0

    End Function

    Public Async Function UpdateLayoutConfigRecord(keyname As String, dgindex As Integer, fieldname As String, fieldvalue As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("layouts")

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' define update document
        Dim doc1 As BsonDocument = New BsonDocument
        With doc1
            .Add("devices." & dgindex & "." & fieldname, fieldvalue)
        End With

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$set", doc1))

        Return 0

    End Function

    Public Async Function UpdateConfigRecord_OLD(keyname As String, dgindex As Integer, fieldarray() As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("fixtures")

        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' define update document
        Dim doc1 As BsonDocument = New BsonDocument
        With doc1
            .Add("configuration." & dgindex & ".startchannel", fieldarray(0))
            .Add("configuration." & dgindex & ".dmxvalue", fieldarray(1))
            .Add("configuration." & dgindex & ".dmxfunctioncategory", fieldarray(2))
            .Add("configuration." & dgindex & ".dmxfunction", fieldarray(3))
        End With

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$set", doc1))

        Return 0

    End Function

    Public Async Function DeleteRecord(keyname As String) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' update value if found
        vCol.DeleteMany(fltr)

        Return 0

    End Function

    Public Async Function DeleteConfigRecord(keyname As String, dgindex As Integer) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("fixtures")


        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' define update document
        Dim doc1 As BsonDocument = New BsonDocument
        With doc1
            .Add("configuration." & dgindex, 1)
        End With

        Dim doc2 As BsonDocument = New BsonDocument
        With doc2
            .Add("configuration", BsonNull.Value)
        End With

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$unset", doc1))
        vCol.UpdateOne(fltr, New BsonDocument("$pull", doc2))

        Return 0

    End Function

    Public Async Function DeleteLayoutDevice(keyname As String, dgindex As Integer) As Task(Of Integer)
        Dim vCol As IMongoCollection(Of BsonDocument)

        vCol = db.GetCollection(Of BsonDocument)("layouts")


        Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of ObjectId)("_id", New ObjectId(keyname))

        ' define update document
        Dim doc1 As BsonDocument = New BsonDocument
        With doc1
            .Add("devices." & dgindex, 1)
        End With

        Dim doc2 As BsonDocument = New BsonDocument
        With doc2
            .Add("devices", BsonNull.Value)
        End With

        ' update value if found
        vCol.UpdateOne(fltr, New BsonDocument("$unset", doc1))
        vCol.UpdateOne(fltr, New BsonDocument("$pull", doc2))

        Return 0

    End Function

    Public Async Function GetResults(fieldname As String, fieldvalue As String) As Task(Of Integer)

        Dim vCol As IMongoCollection(Of BsonDocument)
        vCol = db.GetCollection(Of BsonDocument)("fixtures")

        Dim query As BsonDocument
        query = New BsonDocument("name", fieldvalue)

        Dim myList As List(Of BsonDocument) = Await vCol.Find(query).ToListAsync()

        ' add new item if value not found
        If myList.Count = 0 Then
            ' add record
            Dim emp As BsonDocument = New BsonDocument
            With emp
                .Add(fieldname, fieldvalue)
            End With
            vCol.InsertOne(emp)

        Else
            ' update value if found
            For Each item As BsonDocument In myList
                Dim fltr = Builders(Of BsonDocument).Filter.Eq(Of String)(item.GetElement(0).Name, fieldvalue)
                'MsgBox(item.GetElement(1).Value)
                'GetResults = 1
                'item.Set(fieldname, fieldvalue)
                'vCol.UpdateOne(fltr, New BsonDocument("$set", New BsonDocument(fieldname, fieldvalue)))
            Next
        End If

    End Function

    Public Function UpdateMongoDoc(col As String, nValue As String)
        coll = db.GetCollection(Of BsonDocument)("fixtures")
        Dim fixture As BsonDocument = New BsonDocument
        fixture.Add(col, nValue)

    End Function

End Class

Public Class DeviceConfig
    Public Property startchannel() As String
        Get
            Return m_startchannel
        End Get
        Set
            m_startchannel = Value
        End Set
    End Property
    Private m_startchannel As String
    Public Property dmxvalue() As String
        Get
            Return m_dmxvalue
        End Get
        Set
            m_dmxvalue = Value
        End Set
    End Property
    Private m_dmxvalue As String
    Public Property dmxfunctioncategory() As String
        Get
            Return m_dmxfunctioncategory
        End Get
        Set
            m_dmxfunctioncategory = Value
        End Set
    End Property
    Private m_dmxfunctioncategory As String
    Public Property dmxfunction() As String
        Get
            Return m_dmxfunction
        End Get
        Set
            m_dmxfunction = Value
        End Set
    End Property
    Private m_dmxfunction As String
End Class

Public Class LayoutConfig
    Public Property devicename() As String
        Get
            Return m_devicename
        End Get
        Set
            m_devicename = Value
        End Set
    End Property
    Private m_devicename As String
    Public Property devicealias() As String
        Get
            Return m_devicealias
        End Get
        Set
            m_devicealias = Value
        End Set
    End Property
    Private m_devicealias As String
    Public Property dmxid() As String
        Get
            Return m_dmxid
        End Get
        Set
            m_dmxid = Value
        End Set
    End Property
    Private m_dmxid As String
    Public Property devicegroups() As String
        Get
            Return m_devicegroups
        End Get
        Set
            m_devicegroups = Value
        End Set
    End Property
    Private m_devicegroups As String
End Class


