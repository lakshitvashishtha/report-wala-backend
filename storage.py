import json
import os
from pathlib import Path


class ReportStorage:
    """Persist report jobs and references in MySQL when production settings exist."""

    def __init__(self, local_root):
        self.local_root = Path(local_root)
        requested = os.getenv("STORAGE_BACKEND", "auto").strip().lower()
        file_backend = os.getenv("FILE_STORAGE_BACKEND", "local").strip().lower()
        configured = all(
            primary or railway
            for primary, railway in (
                (os.getenv("MYSQL_HOST"), os.getenv("MYSQLHOST")),
                (os.getenv("MYSQL_DATABASE"), os.getenv("MYSQLDATABASE")),
                (os.getenv("MYSQL_USER"), os.getenv("MYSQLUSER")),
                (os.getenv("MYSQL_PASSWORD"), os.getenv("MYSQLPASSWORD")),
            )
        )
        self.uses_mysql = requested == "mysql" or (requested == "auto" and configured)
        self.files_local = file_backend != "mysql"
        self.mysql_config = {
            "host": os.getenv("MYSQL_HOST") or os.getenv("MYSQLHOST", "127.0.0.1"),
            "port": int(os.getenv("MYSQL_PORT") or os.getenv("MYSQLPORT", "3306")),
            "database": os.getenv("MYSQL_DATABASE") or os.getenv("MYSQLDATABASE", "reportforge"),
            "user": os.getenv("MYSQL_USER") or os.getenv("MYSQLUSER", ""),
            "password": os.getenv("MYSQL_PASSWORD") or os.getenv("MYSQLPASSWORD", ""),
            "charset": "utf8mb4",
        }
        ssl_ca = os.getenv("MYSQL_SSL_CA")
        if ssl_ca:
            self.mysql_config["ssl_ca"] = ssl_ca
            self.mysql_config["ssl_verify_cert"] = True
        if self.files_local:
            self.local_root.mkdir(parents=True, exist_ok=True)
        if self.uses_mysql:
            self._initialize_mysql()
        else:
            self.local_root.mkdir(parents=True, exist_ok=True)

    @property
    def backend_name(self):
        return "mysql" if self.uses_mysql else "local"

    def _connect(self):
        try:
            import mysql.connector
        except ImportError as exc:
            raise RuntimeError(
                "MySQL storage is configured but mysql-connector-python is not installed."
            ) from exc
        return mysql.connector.connect(**self.mysql_config)

    def _initialize_mysql(self):
        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                '''
                CREATE TABLE IF NOT EXISTS users (
                    id CHAR(32) PRIMARY KEY,
                    email VARCHAR(255) UNIQUE NOT NULL,
                    password_hash VARCHAR(255) NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                '''
            )
            cursor.execute(
                '''
                CREATE TABLE IF NOT EXISTS sessions (
                    token CHAR(64) PRIMARY KEY,
                    user_id CHAR(32) NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    expires_at TIMESTAMP NOT NULL,
                    INDEX idx_sessions_user (user_id),
                    CONSTRAINT fk_sessions_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                '''
            )
            
            cursor.execute(
                '''
                CREATE TABLE IF NOT EXISTS report_jobs (
                    id CHAR(32) PRIMARY KEY,
                    user_id CHAR(32) NULL,
                    status VARCHAR(24) NOT NULL,
                    payload LONGTEXT NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                        ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_report_jobs_updated (updated_at)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                '''
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS reference_files (
                    id CHAR(32) PRIMARY KEY,
                    job_id CHAR(32) NOT NULL,
                    filename VARCHAR(255) NOT NULL,
                    content_type VARCHAR(160) NULL,
                    size_bytes BIGINT UNSIGNED NOT NULL,
                    content LONGBLOB NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    INDEX idx_reference_files_job (job_id),
                    CONSTRAINT fk_reference_files_job
                        FOREIGN KEY (job_id) REFERENCES report_jobs(id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS custom_page_files (
                    id CHAR(32) PRIMARY KEY,
                    job_id CHAR(32) NOT NULL,
                    page_role VARCHAR(32) NOT NULL,
                    filename VARCHAR(255) NOT NULL,
                    content_type VARCHAR(160) NULL,
                    size_bytes BIGINT UNSIGNED NOT NULL,
                    content LONGBLOB NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    INDEX idx_custom_page_files_job (job_id),
                    CONSTRAINT fk_custom_page_files_job
                        FOREIGN KEY (job_id) REFERENCES report_jobs(id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS project_detail_files (
                    id CHAR(32) PRIMARY KEY,
                    job_id CHAR(32) NOT NULL,
                    filename VARCHAR(255) NOT NULL,
                    content_type VARCHAR(160) NULL,
                    size_bytes BIGINT UNSIGNED NOT NULL,
                    content LONGBLOB NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    INDEX idx_project_detail_files_job (job_id),
                    CONSTRAINT fk_project_detail_files_job
                        FOREIGN KEY (job_id) REFERENCES report_jobs(id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            try:
                cursor.execute("ALTER TABLE report_jobs ADD COLUMN user_id CHAR(32) NULL")
                cursor.execute("ALTER TABLE report_jobs ADD CONSTRAINT fk_report_jobs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL")
            except:
                pass
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS output_files (
                    id CHAR(32) PRIMARY KEY,
                    job_id CHAR(32) NOT NULL,
                    file_kind VARCHAR(8) NOT NULL,
                    filename VARCHAR(255) NOT NULL,
                    content_type VARCHAR(160) NULL,
                    size_bytes BIGINT UNSIGNED NOT NULL,
                    content LONGBLOB NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE KEY uq_output_job_kind (job_id, file_kind),
                    INDEX idx_output_files_job (job_id),
                    CONSTRAINT fk_output_files_job
                        FOREIGN KEY (job_id) REFERENCES report_jobs(id) ON DELETE CASCADE
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
                """
            )
            connection.commit()
        finally:
            connection.close()

    def save_output_file(self, job_id, file_kind, filename, content):
        """Store a completed report file (docx/pdf) in MySQL for recovery."""
        if not self.uses_mysql:
            return
        import mimetypes as _mimetypes
        content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if file_kind == "docx"
            else "application/pdf"
        )
        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO output_files (id, job_id, file_kind, filename, content_type, size_bytes, content)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE filename = VALUES(filename), content = VALUES(content),
                    size_bytes = VALUES(size_bytes), content_type = VALUES(content_type)
                """,
                (os.urandom(16).hex(), job_id, file_kind, filename, content_type, len(content), content),
            )
            connection.commit()
        finally:
            connection.close()

    def get_output_file(self, job_id, file_kind):
        """Retrieve a completed report file from MySQL. Returns (filename, content) or None."""
        if not self.uses_mysql:
            return None
        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                "SELECT filename, content FROM output_files WHERE job_id = %s AND file_kind = %s",
                (job_id, file_kind),
            )
            row = cursor.fetchone()
            if row:
                return row[0], bytes(row[1])
            return None
        finally:
            connection.close()

    def save_job(self, job_id, payload, user_id=None):
        if not self.uses_mysql:
            return
        serialized = json.dumps(payload, default=str)
        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO report_jobs (id, user_id, status, payload)
                VALUES (%s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE status = VALUES(status), payload = VALUES(payload), user_id = IF(VALUES(user_id) IS NOT NULL, VALUES(user_id), user_id)
                """,
                (job_id, user_id, payload.get("status", "queued"), serialized),
            )
            try:
                cursor.execute("ALTER TABLE report_jobs ADD COLUMN user_id CHAR(32) NULL")
                cursor.execute("ALTER TABLE report_jobs ADD CONSTRAINT fk_report_jobs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL")
            except:
                pass
            connection.commit()
        finally:
            connection.close()

    def load_job(self, job_id):
        if not self.uses_mysql:
            return None
        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute("SELECT payload FROM report_jobs WHERE id = %s", (job_id,))
            row = cursor.fetchone()
            return json.loads(row[0]) if row else None
        finally:
            connection.close()

    def save_reference(self, job_id, filename, content_type, content):
        if self.files_local or not self.uses_mysql:
            job_dir = self.local_root / job_id
            job_dir.mkdir(parents=True, exist_ok=True)
            destination = job_dir / filename
            destination.write_bytes(content)
            return destination

        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO reference_files
                    (id, job_id, filename, content_type, size_bytes, content)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    os.urandom(16).hex(),
                    job_id,
                    filename,
                    content_type,
                    len(content),
                    content,
                ),
            )
            try:
                cursor.execute("ALTER TABLE report_jobs ADD COLUMN user_id CHAR(32) NULL")
                cursor.execute("ALTER TABLE report_jobs ADD CONSTRAINT fk_report_jobs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL")
            except:
                pass
            connection.commit()
        finally:
            connection.close()
        return None

    def materialize_references(self, job_id):
        target = self.local_root / job_id
        if self.files_local or not self.uses_mysql:
            return target if target.exists() and any(target.iterdir()) else None

        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                "SELECT filename, content FROM reference_files WHERE job_id = %s ORDER BY created_at, id",
                (job_id,),
            )
            rows = cursor.fetchall()
        finally:
            connection.close()

        if not rows:
            return None
        target.mkdir(parents=True, exist_ok=True)
        for filename, content in rows:
            (target / filename).write_bytes(bytes(content))
        return target

    def save_custom_page_file(self, job_id, page_role, filename, content_type, content):
        if self.files_local or not self.uses_mysql:
            job_dir = self.local_root / job_id / "_custom_pages"
            job_dir.mkdir(parents=True, exist_ok=True)
            destination = job_dir / f"{page_role}_{filename}"
            destination.write_bytes(content)
            return destination

        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO custom_page_files
                    (id, job_id, page_role, filename, content_type, size_bytes, content)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    os.urandom(16).hex(),
                    job_id,
                    page_role,
                    filename,
                    content_type,
                    len(content),
                    content,
                ),
            )
            try:
                cursor.execute("ALTER TABLE report_jobs ADD COLUMN user_id CHAR(32) NULL")
                cursor.execute("ALTER TABLE report_jobs ADD CONSTRAINT fk_report_jobs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL")
            except:
                pass
            connection.commit()
        finally:
            connection.close()
        return None

    def materialize_custom_page_files(self, job_id):
        target = self.local_root / job_id / "_custom_pages"
        if self.files_local or not self.uses_mysql:
            if not target.exists():
                return {}
            result = {}
            for path in target.iterdir():
                if path.is_file():
                    role = path.name.split("_", 1)[0]
                    result[role] = path
            return result

        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                "SELECT page_role, filename, content FROM custom_page_files WHERE job_id = %s ORDER BY created_at, id",
                (job_id,),
            )
            rows = cursor.fetchall()
        finally:
            connection.close()

        result = {}
        if not rows:
            return result
        target.mkdir(parents=True, exist_ok=True)
        for page_role, filename, content in rows:
            destination = target / f"{page_role}_{filename}"
            destination.write_bytes(bytes(content))
            result[page_role] = destination
        return result

    def save_project_detail_file(self, job_id, filename, content_type, content):
        if self.files_local or not self.uses_mysql:
            job_dir = self.local_root / job_id / "_project_details"
            job_dir.mkdir(parents=True, exist_ok=True)
            destination = job_dir / filename
            destination.write_bytes(content)
            return destination

        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                """
                INSERT INTO project_detail_files
                    (id, job_id, filename, content_type, size_bytes, content)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                (
                    os.urandom(16).hex(),
                    job_id,
                    filename,
                    content_type,
                    len(content),
                    content,
                ),
            )
            try:
                cursor.execute("ALTER TABLE report_jobs ADD COLUMN user_id CHAR(32) NULL")
                cursor.execute("ALTER TABLE report_jobs ADD CONSTRAINT fk_report_jobs_user FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL")
            except:
                pass
            connection.commit()
        finally:
            connection.close()
        return None

    def materialize_project_detail_files(self, job_id):
        target = self.local_root / job_id / "_project_details"
        if self.files_local or not self.uses_mysql:
            return target if target.exists() and any(target.iterdir()) else None

        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute(
                "SELECT filename, content FROM project_detail_files WHERE job_id = %s ORDER BY created_at, id",
                (job_id,),
            )
            rows = cursor.fetchall()
        finally:
            connection.close()

        if not rows:
            return None
        target.mkdir(parents=True, exist_ok=True)
        for filename, content in rows:
            (target / filename).write_bytes(bytes(content))
        return target

    def create_user(self, user_id, email, password_hash):
        if self.uses_mysql:
            connection = self._connect()
            try:
                cursor = connection.cursor()
                cursor.execute("INSERT INTO users (id, email, password_hash) VALUES (%s, %s, %s)", (user_id, email, password_hash))
                connection.commit()
                return True
            except Exception:
                return False
            finally:
                connection.close()
        else:
            # Local file-based fallback
            users_file = self.local_root / "users.json"
            users = {}
            if users_file.exists():
                try:
                    users = json.loads(users_file.read_text())
                except:
                    pass
            users[email] = {"id": user_id, "email": email, "password_hash": password_hash}
            users_file.write_text(json.dumps(users, indent=2))
            return True

    def get_user_by_email(self, email):
        if self.uses_mysql:
            connection = self._connect()
            try:
                cursor = connection.cursor(dictionary=True)
                cursor.execute("SELECT * FROM users WHERE email = %s", (email,))
                return cursor.fetchone()
            finally:
                connection.close()
        else:
            # Local file-based fallback
            users_file = self.local_root / "users.json"
            if users_file.exists():
                try:
                    users = json.loads(users_file.read_text())
                    return users.get(email)
                except:
                    pass
            return None

    def get_user_by_id(self, user_id):
        if not self.uses_mysql: return None
        connection = self._connect()
        try:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("SELECT * FROM users WHERE id = %s", (user_id,))
            return cursor.fetchone()
        finally:
            connection.close()

    def create_session(self, token, user_id, expires_at):
        if self.uses_mysql:
            connection = self._connect()
            try:
                cursor = connection.cursor()
                cursor.execute("INSERT INTO sessions (token, user_id, expires_at) VALUES (%s, %s, %s)", (token, user_id, expires_at))
                connection.commit()
                return True
            finally:
                connection.close()
        else:
            # Local file-based fallback
            sessions_file = self.local_root / "sessions.json"
            sessions = {}
            if sessions_file.exists():
                try:
                    sessions = json.loads(sessions_file.read_text())
                except:
                    pass
            sessions[token] = {"user_id": user_id, "expires_at": expires_at}
            sessions_file.write_text(json.dumps(sessions, indent=2))
            return True

    def get_session(self, token):
        if self.uses_mysql:
            connection = self._connect()
            try:
                cursor = connection.cursor(dictionary=True)
                cursor.execute("SELECT user_id, expires_at FROM sessions WHERE token = %s AND expires_at > NOW()", (token,))
                return cursor.fetchone()
            finally:
                connection.close()
        else:
            # Local file-based fallback
            sessions_file = self.local_root / "sessions.json"
            if sessions_file.exists():
                try:
                    sessions = json.loads(sessions_file.read_text())
                    session = sessions.get(token)
                    if session:
                        # Check if session is expired
                        from datetime import datetime
                        expires_at = datetime.strptime(session["expires_at"], "%Y-%m-%d %H:%M:%S")
                        if expires_at > datetime.utcnow():
                            return session
                except:
                    pass
            return None

    def get_user_jobs(self, user_id):
        if self.uses_mysql:
            connection = self._connect()
            try:
                cursor = connection.cursor()
                # Only return jobs from the last 24 hours for download recovery
                cursor.execute(
                    """SELECT id, payload, created_at FROM report_jobs
                       WHERE user_id = %s AND created_at > NOW() - INTERVAL 24 HOUR
                       ORDER BY created_at DESC LIMIT 20""",
                    (user_id,)
                )
                rows = cursor.fetchall()
                import json
                return [{"id": r[0], "payload": json.loads(r[1]), "created_at": r[2].isoformat() if r[2] else None} for r in rows]
            finally:
                connection.close()
        else:
            # Local file-based fallback - return empty list for now
            return []

    def cleanup_old_jobs(self, days=7):
        if not self.uses_mysql: return
        connection = self._connect()
        try:
            cursor = connection.cursor()
            cursor.execute("DELETE FROM report_jobs WHERE created_at < NOW() - INTERVAL %s DAY", (days,))
            connection.commit()
            return cursor.rowcount
        finally:
            connection.close()
