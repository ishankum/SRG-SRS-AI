using System;
using System.Collections.Generic;
using System.Data;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;

namespace SrsAi.Functions
{
    public class DatabaseSchemaReader
    {
        private readonly ILogger _logger;

        public DatabaseSchemaReader(ILogger logger)
        {
            _logger = logger;
        }

        public async Task<string> ExtractSchemaJsonAsync(List<string> connectionStrings)
        {
            var databaseSchemas = new List<object>();

            foreach (string connStr in connectionStrings)
            {
                if (string.IsNullOrWhiteSpace(connStr)) continue;

                try
                {
                    var builder = new SqlConnectionStringBuilder(connStr);
                    string catalogName = builder.InitialCatalog;
                    _logger.LogInformation($"Extracting DB schema for catalog: '{catalogName}'");

                    using var conn = new SqlConnection(connStr);
                    await conn.OpenAsync();

                    // 1. Extract Tables and Columns
                    var tablesList = new List<object>();
                    string tablesQuery = @"
                        SELECT TABLE_NAME 
                        FROM INFORMATION_SCHEMA.TABLES 
                        WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME NOT LIKE 'sys%'
                        ORDER BY TABLE_NAME";

                    using (var tableCmd = new SqlCommand(tablesQuery, conn))
                    using (var tableReader = await tableCmd.ExecuteReaderAsync())
                    {
                        var tableNames = new List<string>();
                        while (await tableReader.ReadAsync())
                        {
                            tableNames.Add(tableReader.GetString(0));
                        }
                        await tableReader.CloseAsync();

                        foreach (string tableName in tableNames)
                        {
                            var columnsList = new List<object>();
                            string columnsQuery = @"
                                SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
                                FROM INFORMATION_SCHEMA.COLUMNS
                                WHERE TABLE_NAME = @TableName
                                ORDER BY ORDINAL_POSITION";

                            using var colCmd = new SqlCommand(columnsQuery, conn);
                            colCmd.Parameters.AddWithValue("@TableName", tableName);

                            using var colReader = await colCmd.ExecuteReaderAsync();
                            while (await colReader.ReadAsync())
                            {
                                string colName = colReader.GetString(0);
                                string dataType = colReader.GetString(1);
                                object maxLen = colReader.IsDBNull(2) ? "N/A" : colReader.GetValue(2);
                                string isNullable = colReader.GetString(3);

                                columnsList.Add(new
                                {
                                    name = colName,
                                    type = dataType,
                                    maxLength = maxLen.ToString(),
                                    isNullable = isNullable
                                });
                            }
                            await colReader.CloseAsync();

                            tablesList.Add(new
                            {
                                tableName = tableName,
                                columnsCount = columnsList.Count,
                                columns = columnsList
                            });
                        }
                    }

                    // 2. Extract Stored Procedures
                    var procList = new List<string>();
                    string procQuery = @"
                        SELECT ROUTINE_NAME 
                        FROM INFORMATION_SCHEMA.ROUTINES 
                        WHERE ROUTINE_TYPE = 'PROCEDURE' AND ROUTINE_NAME NOT LIKE 'sp_%'
                        ORDER BY ROUTINE_NAME";

                    using (var procCmd = new SqlCommand(procQuery, conn))
                    using (var procReader = await procCmd.ExecuteReaderAsync())
                    {
                        while (await procReader.ReadAsync())
                        {
                            procList.Add(procReader.GetString(0));
                        }
                    }

                    databaseSchemas.Add(new
                    {
                        catalog = catalogName,
                        totalTables = tablesList.Count,
                        totalStoredProcedures = procList.Count,
                        tables = tablesList,
                        storedProcedures = procList
                    });

                    _logger.LogInformation($"Successfully extracted {tablesList.Count} tables and {procList.Count} procedures for catalog '{catalogName}'.");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Failed to extract DB schema for connection string: {ex.Message}");
                }
            }

            var rootResult = new
            {
                generatedAtUtc = DateTime.UtcNow,
                totalCatalogs = databaseSchemas.Count,
                databases = databaseSchemas
            };

            return JsonSerializer.Serialize(rootResult, new JsonSerializerOptions { WriteIndented = true });
        }
    }
}
