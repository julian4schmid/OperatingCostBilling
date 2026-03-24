import subprocess
from datetime import datetime
import os

from config import (
    PG_DUMP_PATH,
    DB_HOST,
    DB_PORT,
    DB_NAME,
    DB_USER,
    DB_PASSWORD,
    BACKUP_DIR
)


def create_backup():
    # Ensure backup directory exists
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    # Create timestamp for unique file naming
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    backup_file = BACKUP_DIR / f"{DB_NAME}_backup_{timestamp}.sql"

    # Build pg_dump command
    command = [
        PG_DUMP_PATH,
        "-h", DB_HOST,
        "-p", str(DB_PORT),
        "-U", DB_USER,
        "-F", "p",  # plain SQL format
        "-f", str(backup_file),
        DB_NAME
    ]

    # Set password via environment variable (avoids prompt)
    env = os.environ.copy()
    env["PGPASSWORD"] = DB_PASSWORD

    try:
        print(f"Starting backup: {backup_file}")

        # Execute pg_dump command
        subprocess.run(command, env=env, check=True)

        print("Backup created successfully")

    except subprocess.CalledProcessError as e:
        print("Backup failed:", e)


if __name__ == "__main__":
    create_backup()
