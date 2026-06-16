import json
import os
import sqlite3

TARGET_SN = "260519T2P8W04S"
DB_PATH = os.path.join(os.environ["LOCALAPPDATA"], "PaperbellAppDotNet", "app.db")


def norm(value: str | None) -> str:
    return (value or "").strip().replace(" ", "").lower()


def key_model_item(model_sku: str | None, item_sku: str | None) -> str:
    return norm(model_sku) + norm(item_sku)


def line_id(item: dict, idx: int) -> str:
    item_id = item.get("item_id") or 0
    model_id = item.get("model_id") or 0
    if model_id:
        return f"{item_id}:{model_id}"

    order_item_id = item.get("order_item_id")
    if order_item_id not in (None, "", 0):
        return str(order_item_id)

    model_name = item.get("model_name") or item.get("variation_name") or ""
    if item_id and model_name:
        return f"{item_id}:n:{norm(model_name)}"

    return f"line:{idx}"


def migrate_if_needed(con: sqlite3.Connection) -> None:
    cur = con.cursor()
    cols = [r[1] for r in cur.execute("PRAGMA table_info(order_process)")]
    if "order_item_id" in cols:
        return

    cur.executescript(
        """
        BEGIN;
        CREATE TABLE order_process_new (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            order_sn TEXT,
            order_item_id TEXT NOT NULL,
            item_key TEXT,
            model_sku TEXT,
            item_sku TEXT,
            item_name TEXT,
            model_name TEXT,
            qty INTEGER,
            status TEXT,
            create_time INTEGER,
            saved_at INTEGER,
            printed INTEGER DEFAULT 0,
            printed_at INTEGER,
            UNIQUE(order_sn, order_item_id)
        );

        INSERT INTO order_process_new(
            id, order_sn, order_item_id, item_key, model_sku, item_sku, item_name, model_name,
            qty, status, create_time, saved_at, printed, printed_at
        )
        SELECT
            id, order_sn, 'legacy:' || id, item_key, model_sku, item_sku, item_name, model_name,
            qty, status, create_time, saved_at, printed, printed_at
        FROM order_process;

        DROP TABLE order_process;
        ALTER TABLE order_process_new RENAME TO order_process;
        COMMIT;
        """
    )
    con.commit()


def rebuild_from_raw_json(con: sqlite3.Connection) -> None:
    cur = con.cursor()
    rows = list(
        cur.execute(
            "SELECT order_sn, raw_json FROM orders WHERE raw_json IS NOT NULL AND raw_json <> ''"
        )
    )

    for order_sn, raw in rows:
        try:
            order = json.loads(raw)
        except Exception:
            continue

        items = order.get("item_list") or []
        if not items:
            continue

        status = order.get("order_status") or ""
        create_time = int(order.get("create_time") or 0)
        keep_ids: list[str] = []

        for idx, item in enumerate(items):
            item_sku = item.get("item_sku") or ""
            model_sku = (item.get("model_sku") or item.get("variation_sku") or "").strip()
            item_name = item.get("item_name") or ""
            model_name = item.get("model_name") or ""
            qty = int(item.get("model_quantity_purchased") or 1)

            oid = line_id(item, idx)
            keep_ids.append(oid)

            cur.execute(
                """
                INSERT INTO order_process(
                    order_sn, order_item_id, item_key, model_sku, item_sku, item_name, model_name,
                    qty, status, create_time, saved_at
                )
                VALUES(?,?,?,?,?,?,?,?,?,?,strftime('%s','now'))
                ON CONFLICT(order_sn, order_item_id) DO UPDATE SET
                    item_key=excluded.item_key,
                    model_sku=excluded.model_sku,
                    item_sku=excluded.item_sku,
                    item_name=excluded.item_name,
                    model_name=excluded.model_name,
                    qty=excluded.qty,
                    status=excluded.status,
                    create_time=excluded.create_time,
                    saved_at=excluded.saved_at
                """,
                (
                    order_sn,
                    oid,
                    key_model_item(model_sku, item_sku),
                    model_sku,
                    item_sku,
                    item_name,
                    model_name,
                    qty,
                    status,
                    create_time,
                ),
            )

        if keep_ids:
            placeholders = ",".join("?" for _ in keep_ids)
            cur.execute(
                f"DELETE FROM order_process WHERE order_sn = ? AND order_item_id NOT IN ({placeholders})",
                [order_sn, *keep_ids],
            )

    con.commit()


def print_target(con: sqlite3.Connection) -> None:
    cur = con.cursor()
    cols = [r[1] for r in cur.execute("PRAGMA table_info(order_process)")]
    print("order_process columns:", cols)
    rows = list(
        cur.execute(
            """
            SELECT id, order_item_id, item_key, model_sku, item_sku, model_name, qty
            FROM order_process
            WHERE order_sn = ?
            ORDER BY id
            """,
            (TARGET_SN,),
        )
    )
    print("target rows:", len(rows))
    for row in rows:
        print(row)


def main() -> None:
    print("db:", DB_PATH, "exists:", os.path.exists(DB_PATH))
    if not os.path.exists(DB_PATH):
        return

    con = sqlite3.connect(DB_PATH)
    try:
        migrate_if_needed(con)
        rebuild_from_raw_json(con)
        print_target(con)
    finally:
        con.close()


if __name__ == "__main__":
    main()
