# AccessData

A simple helper class for working with Microsoft Access databases using OleDb.

This library provides a minimal abstraction over common database operations, allowing execution of queries without repeating standard connection and command logic.

## Functionality

* Execution of SELECT queries
* Execution of INSERT, UPDATE, and DELETE operations
* Retrieval of last inserted ID
* Support for parameterized queries

## Example

```csharp
var db = new AccessDatabase("path_to_database.accdb");

db.ExecuteQuery("SELECT * FROM Users");

if (!db.HasError)
{
    var table = db.Table;
}
```

## Notes

* Uses OleDb provider (Microsoft Access)
* Intended for Windows environments
* Designed for simple scenarios, not as a full data access layer

---

## Additional note

This project was created as a lightweight utility for everyday use when working with Access-based systems.
