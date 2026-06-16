"""Tandai order dibuat Maret ke bawah (sebelum 1 Apr 2026 WIB) sebagai sudah dibungkus."""
import os
import sqlite3
import time
from datetime import datetime, timezone, timedelta

DB_PATH = os.path.join(os.environ["LOCALAPPDATA"], "PaperbellAppDotNet", "app.db")
STATE_KEY = "backfill_through_march2026_packaged_v1"

# 1 April 2026 00:00 WIB
WIB = timezone(timedelta(hours=7))
CUTOFF_UNIX = int(datetime(2026, 4, 1, 0, 0, 0, tzinfo=WIB).timestamp())


def main() -> None:
    if not os.path.isfile(DB_PATH):
        print(f"DB tidak ditemukan: {DB_PATH}")
        return

    con = sqlite3.connect(DB_PATH)
    try:
        cur = con.cursor()
        cols = [r[1] for r in cur.execute("PRAGMA table_info(orders)")]
        if "packaged" not in cols:
            print("Kolom orders.packaged belum ada â€” jalankan aplikasi sekali dulu.")
            return

        row = cur.execute(
            "SELECT value FROM app_state WHERE key = ? LIMIT 1", (STATE_KEY,)
        ).fetchone()
        if row and row[0] == "1":
            print("Backfill sudah pernah dijalankan (app_state).")
            print(f"Hapus baris app_state '{STATE_KEY}' jika ingin ulang.")
            return

        now = int(time.time())
        orders = cur.execute(
            """
            SELECT op.order_sn, MAX(op.create_time)
            FROM order_process op
            WHERE op.create_time < ?
              AND (IFNULL(UPPER(TRIM(op.status)), '') <> 'CANCELLED')
            GROUP BY op.order_sn
            """,
            (CUTOFF_UNIX,),
        ).fetchall()

        n = 0
        for order_sn, create_time in orders:
            cur.execute(
                """
                UPDATE orders
                SET packaged = 1,
                    packaged_at = COALESCE(packaged_at, ?),
                    update_time = ?
                WHERE order_sn = ?
                """,
                (now, now, order_sn),
            )
            if cur.rowcount == 0:
                cur.execute(
                    """
                    INSERT INTO orders(order_sn, packaged, packaged_at, create_time, update_time)
                    VALUES (?, 1, ?, ?, ?)
                    """,
                    (order_sn, now, create_time or now, now),
                )
            n += 1

        cur.execute(
            "INSERT OR REPLACE INTO app_state(key, value) VALUES(?, '1')",
            (STATE_KEY,),
        )
        con.commit()
        print(f"Selesai: {n} order (dibuat sebelum 1 Apr 2026 WIB) ditandai sudah dibungkus.")
    finally:
        con.close()


if __name__ == "__main__":
    main()
