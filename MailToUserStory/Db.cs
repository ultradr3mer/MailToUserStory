
using Microsoft.Data.Sqlite;
// ----------------------------
// DB layer (SQLite)
// ----------------------------
sealed class Db : IDisposable
{
  private readonly SqliteConnection _conn;
  public Db(string path)
  {
    _conn = new SqliteConnection("Data Source=" + path);
    _conn.Open();
  }
  public SqliteConnection Connection => _conn;
  public void Dispose() => _conn.Dispose();

  public string? GetDeltaLink(string mailbox)
  {
    using var cmd = _conn.CreateCommand();
    cmd.CommandText = "SELECT delta_link FROM Mailboxes WHERE address=@a LIMIT 1";
    cmd.Parameters.AddWithValue("@a", mailbox);
    return cmd.ExecuteScalar() as string;
  }

  public void UpsertDeltaLink(string mailbox, string? delta)
  {
    using var tx = _conn.BeginTransaction();
    using var cmd = _conn.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = @"
INSERT INTO Mailboxes(address, delta_link) VALUES(@a, @d)
ON CONFLICT(address) DO UPDATE SET delta_link=excluded.delta_link";
    cmd.Parameters.AddWithValue("@a", mailbox);
    cmd.Parameters.AddWithValue("@d", (object?)delta ?? DBNull.Value);
    cmd.ExecuteNonQuery();
    tx.Commit();
  }

  public bool WasProcessed(string messageId)
  {
    using var cmd = _conn.CreateCommand();
    cmd.CommandText = "SELECT 1 FROM ProcessedEmails WHERE graph_message_id=@id LIMIT 1";
    cmd.Parameters.AddWithValue("@id", messageId);
    using var r = cmd.ExecuteReader();
    return r.Read();
  }

  public void MarkProcessed(string messageId, string mailbox, int? workItemId, string outcome)
  {
    using var tx = _conn.BeginTransaction();
    using var cmd = _conn.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = @"
INSERT INTO ProcessedEmails(graph_message_id, mailbox, work_item_id, processed_at, outcome)
VALUES(@id, @mb, @wi, @ts, @out)";
    cmd.Parameters.AddWithValue("@id", messageId);
    cmd.Parameters.AddWithValue("@mb", mailbox);
    cmd.Parameters.AddWithValue("@wi", (object?)workItemId ?? DBNull.Value);
    cmd.Parameters.AddWithValue("@ts", DateTimeOffset.UtcNow.ToString("O"));
    cmd.Parameters.AddWithValue("@out", outcome);
    cmd.ExecuteNonQuery();
    tx.Commit();
  }

  public void LinkStory(string mailbox, int workItemId)
  {
    using var tx = _conn.BeginTransaction();
    using var cmd = _conn.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = @"
INSERT OR IGNORE INTO Stories(work_item_id, mailbox) VALUES(@wi, @mb)";
    cmd.Parameters.AddWithValue("@wi", workItemId);
    cmd.Parameters.AddWithValue("@mb", mailbox);
    cmd.ExecuteNonQuery();
    tx.Commit();
  }

  public static void InitializeSchema(Db db)
  {
    using var cmd = db.Connection.CreateCommand();
    cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS Mailboxes(
  address TEXT PRIMARY KEY,
  delta_link TEXT
);
CREATE TABLE IF NOT EXISTS Stories(
  work_item_id INTEGER,
  mailbox TEXT,
  PRIMARY KEY(work_item_id, mailbox)
);
CREATE TABLE IF NOT EXISTS ProcessedEmails(
  graph_message_id TEXT PRIMARY KEY,
  mailbox TEXT NOT NULL,
  work_item_id INTEGER NULL,
  processed_at TEXT NOT NULL,
  outcome TEXT NOT NULL
);
";
    cmd.ExecuteNonQuery();
  }
}
