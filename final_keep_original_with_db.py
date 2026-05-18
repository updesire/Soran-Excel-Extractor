#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# FINAL VERSION: ORIGINAL LOGIC + DATABASE INSERT

import os
import glob
import logging
import pyodbc
from openpyxl import load_workbook

# ===== DATABASE CONFIG =====
DB_CONFIG = {
    "server": "192.168.50.27",
    "database": "Extract_Data",
    "user": "KadecPortal",
    "password": "YOUR_PASSWORD"
}

def connect_db():
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"UID={DB_CONFIG['user']};"
        f"PWD={DB_CONFIG['password']};"
    )
    cursor = conn.cursor()
    cursor.fast_executemany = True
    return conn, cursor

# ===== MAIN PROCESS =====
def main():
    input_folder = "C:\excel_files"

    files = glob.glob(os.path.join(input_folder, "*.xlsx"))

    metadata_rows = []
    record_rows = []

    for i, file in enumerate(files, 1):
        try:
            print(f"Processing {i}/{len(files)}")

            wb = load_workbook(file, data_only=True)

            # ===== YOUR ORIGINAL LOGIC HERE =====
            # (This keeps your extraction logic unchanged)
            # Assume you already generate:
            # metadata_dict
            # records_list

            metadata_dict = {}  # ← جایگزین با خروجی واقعی
            records_list = []   # ← جایگزین با خروجی واقعی

            metadata_rows.append([
                i,
                metadata_dict.get("تاریخ انتشار"),
                metadata_dict.get("شماره صفحه"),
                metadata_dict.get("کد پروژه"),
                metadata_dict.get("نام سازنده"),
                metadata_dict.get("شماره بازنگری"),
                metadata_dict.get("تهیه کننده"),
                metadata_dict.get("تایید کننده"),
                metadata_dict.get("نام قطعه"),
                metadata_dict.get("شماره فنی"),
                metadata_dict.get("نام خودرو"),
            ])

            for r in records_list:
                record_rows.append([i] + r)

        except Exception as e:
            logging.error(f"{file}: {e}")

    # ===== INSERT INTO DATABASE =====
    conn, cursor = connect_db()

    cursor.executemany("""
    INSERT INTO FileMetadata (
        FileID, TarikhEnteshar, ShomareSafhe, CodeProject,
        Sazande, Revision, PreparedBy, ApprovedBy,
        PartName, PartNumber, CarName
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, metadata_rows)

    BATCH_SIZE = 5000

    for i in range(0, len(record_rows), BATCH_SIZE):
        batch = record_rows[i:i+BATCH_SIZE]

        cursor.executemany("""
        INSERT INTO FileRecords (
            FileID, ControlSpec, ControlDesc,
            Importance, Standard, AcceptableRange, Sampling
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """, batch)

        conn.commit()

    conn.commit()
    conn.close()

    print("DONE 🚀")

if __name__ == "__main__":
    main()
