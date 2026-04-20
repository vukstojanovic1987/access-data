using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Runtime.Versioning;

namespace AccessData
{
    /// <summary>
    /// Provides basic data access functionality for Microsoft Access databases using OleDb.
    /// Supports query execution, non-query commands, and insert operations with identity retrieval.
    /// </summary>
    [SupportedOSPlatform("windows")]
    public class AccessDatabase : IDisposable
    {
        private readonly List<OleDbParameter> _parameters = [];
        private DataTable? _table;
        private string _errorMessage = string.Empty;
        private bool _disposed;
        private readonly string _databasePath;

        /// <summary>
        /// Initializes a new instance of the <see cref="AccessDatabase"/> class.
        /// </summary>
        /// <param name="databasePath">Full path to the Access database file.</param>
        public AccessDatabase(string databasePath) => _databasePath = databasePath;

        /// <summary>
        /// Gets the result of the last executed query.
        /// </summary>
        public DataTable? Table => _table;

        /// <summary>
        /// Gets the last error message.
        /// </summary>
        public string ErrorMessage => _errorMessage;

        /// <summary>
        /// Indicates whether the last operation resulted in an error.
        /// </summary>
        public bool HasError => !string.IsNullOrWhiteSpace(_errorMessage);

        /// <summary>
        /// Executes a SELECT query and stores the result in a DataTable.
        /// </summary>
        /// <param name="query">SQL query string.</param>
        public void ExecuteQuery(string query)
        {
            _errorMessage = string.Empty;
            _table = new DataTable();

            using var connection = CreateConnection();

            try
            {
                connection.Open();

                using var command = new OleDbCommand(query, connection);
                AddParameters(command);

                using var adapter = new OleDbDataAdapter(command);
                adapter.Fill(_table);
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Executes INSERT, UPDATE, or DELETE commands.
        /// </summary>
        /// <param name="query">SQL query string.</param>
        /// <returns>Number of affected rows.</returns>
        public int ExecuteNonQuery(string query)
        {
            _errorMessage = string.Empty;

            using var connection = CreateConnection();

            try
            {
                connection.Open();

                using var command = new OleDbCommand(query, connection);
                AddParameters(command);

                return command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                return 0;
            }
        }

        /// <summary>
        /// Executes an INSERT command and returns the newly generated ID.
        /// </summary>
        /// <param name="insertQuery">SQL INSERT query.</param>
        /// <returns>Newly created record ID or -1 if failed.</returns>
        public int ExecuteInsertAndGetId(string insertQuery)
        {
            _errorMessage = string.Empty;

            using var connection = CreateConnection();
            connection.Open();

            using var transaction = connection.BeginTransaction();

            try
            {
                using var insertCommand = new OleDbCommand(insertQuery, connection, transaction);
                AddParameters(insertCommand);

                insertCommand.ExecuteNonQuery();

                using var identityCommand = new OleDbCommand("SELECT @@IDENTITY;", connection, transaction);
                object? result = identityCommand.ExecuteScalar();

                transaction.Commit();

                if (result != null && int.TryParse(result.ToString(), out int id))
                    return id;

                return -1;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                _errorMessage = ex.Message;
                return -1;
            }
        }

        /// <summary>
        /// Adds a parameter to the current query.
        /// </summary>
        /// <param name="name">Parameter name.</param>
        /// <param name="value">Parameter value.</param>
        public void AddParameter(string name, object value)
        {
            _parameters.Add(new OleDbParameter(name, value));
        }

        /// <summary>
        /// Clears all parameters.
        /// </summary>
        public void ClearParameters()
        {
            _parameters.Clear();
        }

        private OleDbConnection CreateConnection()
        {
            return new OleDbConnection(
                $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath}");
        }

        private void AddParameters(OleDbCommand command)
        {
            foreach (var param in _parameters)
            {
                command.Parameters.Add(param);
            }

            _parameters.Clear();
        }

        /// <summary>
        /// Releases resources used by the class.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;

            _parameters.Clear();
            _table?.Dispose();
            _table = null;

            _disposed = true;
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Finalizer.
        /// </summary>
        ~AccessDatabase()
        {
            Dispose();
        }
    }
}