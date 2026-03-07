import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

conn = psycopg2.connect(dbname="postgres", user="postgres", password="postgres", host="127.0.0.1", port=5432)
conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
cur = conn.cursor()
cur.execute("SELECT 1 FROM pg_database WHERE datname='summary_report'")
if cur.fetchone():
    print("Database summary_report already exists")
else:
    cur.execute("CREATE DATABASE summary_report")
    print("Database summary_report created")
cur.close()
conn.close()
