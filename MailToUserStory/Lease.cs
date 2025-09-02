// ----------------------------
// Single-instance lease
// ----------------------------
sealed class Lease : IDisposable
{
  private readonly Db _db;
  private bool _released;
  private Lease(Db db) { _db = db; }

  public static async Task<Lease> AcquireAsync(Db db, TimeSpan duration)
  {
    string owner = Environment.MachineName + ":" + Environment.ProcessId;
    while (true)
    {
      using var tx = db.Connection.BeginTransaction();
      using var cmdSel = db.Connection.CreateCommand();
      cmdSel.Transaction = tx;
      cmdSel.CommandText = "SELECT owner, expires_at FROM Lease WHERE id=1";
      using var r = cmdSel.ExecuteReader();
      string? curOwner = null; DateTimeOffset? expires = null;
      if (r.Read())
      {
        curOwner = r.IsDBNull(0) ? null : r.GetString(0);
        expires = r.IsDBNull(1) ? null : DateTimeOffset.Parse(r.GetString(1));
      }
      r.Close();

      bool canTake = curOwner == null || expires == null || expires < DateTimeOffset.UtcNow;
      using var cmd = db.Connection.CreateCommand();
      cmd.Transaction = tx;
      if (canTake)
      {
        cmd.CommandText = @"
INSERT INTO Lease(id, owner, expires_at) VALUES(1, @o, @e)
ON CONFLICT(id) DO UPDATE SET owner=excluded.owner, expires_at=excluded.expires_at";
        cmd.Parameters.AddWithValue("@o", owner);
        cmd.Parameters.AddWithValue("@e", DateTimeOffset.UtcNow.Add(duration).ToString("O"));
        cmd.ExecuteNonQuery();
        tx.Commit();
        return new Lease(db);
      }
      tx.Rollback();
      await Task.Delay(1000);
    }
  }

  public void Dispose()
  {
    if (_released) return;
    using var tx = _db.Connection.BeginTransaction();
    using var cmd = _db.Connection.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = "UPDATE Lease SET owner=NULL, expires_at=NULL WHERE id=1";
    cmd.ExecuteNonQuery();
    tx.Commit();
    _released = true;
  }
}
