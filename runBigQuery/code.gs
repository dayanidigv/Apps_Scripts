// Click "Add a Service" to enable the BigQuery API

function runBigQuery() {
  var projectId = 'your-project-id';  // Replace with your Google Cloud project ID
  var query = `
    WITH d_phoneno AS (
      SELECT phoneno, COUNT(*) 
      FROM \`your-project-id.your-dataset-id.your-table-id\`
      WHERE phoneno != 'null'
      GROUP BY phoneno
      HAVING COUNT(*) > 1
    )
    SELECT name, phoneno
    FROM \`your-project-id.your-dataset-id.your-table-id\`
    WHERE phoneno NOT IN (SELECT phoneno FROM d_phoneno);
  `;

  var request = {
    query: query,
    useLegacySql: false,
  };

  // Run the query using BigQuery API
  var queryResults = BigQuery.Jobs.query(request, projectId);

  // Check for errors in query results
  if (queryResults.error) {
    Logger.log('Error: ' + queryResults.error.message);
    return;
  }

  // Parse and log the results
  var rows = queryResults.rows;
  if (rows && rows.length > 0) {
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i].f;
      var name = row[0].v;
      var phoneno = row[1].v;
      Logger.log('Name: ' + name + ', Phone Number: ' + phoneno);
    }
  } else {
    Logger.log('No results found.');
  }
}
