# MS Access Utilities

## What is this?
I recently completed some technical analysis on a Microsoft Access database that was critical to the customers operations and was still actively being developed.

This collection of VBS modules was developed in response to some of the findings, so that the community can benefit from some simple utility modules.

## Installation

To install any of the modules e.g. SqlServerDataProvider, you need to import the `.cls` file into your MS Access project:

1. Open your MS Access database.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, go to `File -> Import File...`.
4. Select the `SqlServerDataProvider.cls` file and click `Open`.
5. Ensure you have added a reference to Microsoft ActiveX Data Objects 6.1 Library and Microsoft Scripting Runtime.

.cls was selected as the format for this repo because it requires the least administrative burden for MS Access developers.

Alternatives:
* C++ DLL - requires regsvr32 registration that requires elevated privleges often disabled on CyberSecurity conscious organisations.
* .NET DLL - requires the .NET runtime to be installed on the host machine, which requires elevated priviledges. Or, bundling the .NET runtime with the assembly, bloating the library footprint.

However, if enough community support of a DLL is raised, they can be developed on a further version.

## Usage

### Initialization
```vb
Option Explicit

Dim clsProvider As SqlServerDataProvider

Set clsProvider = New SqlServerDataProvider

clsProvider.Initialize "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=test;User ID=test;Password=test;", 0, 30

' Code removed for brevity...
```

### Execute a Query (no resultset)
```vb
Option Explicit

Dim clsProvider As SqlServerDataProvider
Dim dictParams As Scripting.Dictionary

'Initialisation code removed for brevity...

Set dictParams = New Scripting.Dictionary
dictParams.Add "p1", 1
dictParams.Add "p2", "Two"

clsProvider.ExecuteNonQuery("UPDATE dbo.test1 SET Name=@p2 WHERE ID=@p1;",dictParams)
```

### Execute a Scalar query (single value)
```vb
Option Explicit

Dim clsProvider As SqlServerDataProvider
Dim dictParams As Scripting.Dictionary
Dim returnValue As Variant

'Initialisation code removed for brevity...

Set dictParams = New Scripting.Dictionary
dictParams.Add "p1", 1

returnValue = clsProvider.ExecuteScalar("SELECT LastUpdateOn FROM dbo.test1 WHERE ID=@p1;",dictParams)
```

### Execute a Recordset query (DataTable resultset)
```vb
Option Explicit

Dim clsProvider As SqlServerDataProvider
Dim dictParams As Scripting.Dictionary
Dim rs As ADODB.Recordset

'Initialisation code removed for brevity...

Set dictParams = New Scripting.Dictionary
dictParams.Add "p1", 1

Set rs = clsProvider.GetRecordset("SELECT * FROM dbo.test1 WHERE ID=@p1;",dictParams)
rs.Open

'...

clsProvider.CleanUpRecordset rs

```

### Reusing connections
If your application needs to keep an active connection open to the database consider using
the additional methods of "OnConnection" to reuse the same connection object. This prevents
creating a new connection object from the pool for each query to the database, an example is shown below:

```vb
Option Explicit

Dim clsProvider As SqlServerDataProvider
Dim dictParams As Scripting.Dictionary
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection

'Initialisation code removed for brevity...

Set dictParams = New Scripting.Dictionary
dictParams.Add "p1", 1

Set conn = clsProvider.GetDbConnection()

'... Run multiple statements passing in the connection.

'Pass the active connection into GetRecordset using its overload method.
Set rs = clsProvider.GetRecordsetOnConnection(conn, "SELECT * FROM dbo.test1 WHERE ID=@p1;",dictParams)
rs.Open

'...

clsProvider.CleanUpRecordset rs

```

## Contributing
We welcome contributions to this

 project! If you have an idea for a new feature or have found a bug, please open an issue on GitHub. If you would like to contribute code, feel free to fork the repository and submit a pull request. Make sure to follow our coding guidelines and include tests for any new features or bug fixes.

## License

This project is licensed under the MIT License.
