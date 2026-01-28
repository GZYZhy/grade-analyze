import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

DB_PATH = Path(__file__).resolve().parent / "data" / "grade_analyze.db"


@dataclass
class ClassItem:
    id: int
    name: str


@dataclass
class StudentItem:
    id: int
    name: str
    class_id: int


@dataclass
class ExamItem:
    id: int
    name: str
    class_id: int


class DataStore:
    def __init__(self, db_path: Path = DB_PATH):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS classes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE
                );

                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    class_id INTEGER NOT NULL,
                    UNIQUE(name, class_id),
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS exams (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    class_id INTEGER NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(name, class_id),
                    FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS scores (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    exam_id INTEGER NOT NULL,
                    subject TEXT NOT NULL,
                    score REAL,
                    score_raw REAL,
                    rank INTEGER,
                    total_score REAL,
                    total_raw REAL,
                    grade_rank INTEGER,
                    class_rank INTEGER,
                    FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
                    FOREIGN KEY(exam_id) REFERENCES exams(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                );
                """
            )
            columns = [row["name"] for row in conn.execute("PRAGMA table_info(scores)").fetchall()]
            if "score_raw" not in columns:
                conn.execute("ALTER TABLE scores ADD COLUMN score_raw REAL")

    def get_classes(self) -> List[ClassItem]:
        with self._connect() as conn:
            rows = conn.execute("SELECT id, name FROM classes ORDER BY id").fetchall()
        return [ClassItem(id=row["id"], name=row["name"]) for row in rows]

    def add_class(self, name: str) -> ClassItem:
        with self._connect() as conn:
            conn.execute("INSERT OR IGNORE INTO classes(name) VALUES(?)", (name,))
            row = conn.execute("SELECT id, name FROM classes WHERE name=?", (name,)).fetchone()
        return ClassItem(id=row["id"], name=row["name"])

    def get_students(self, class_id: int) -> List[StudentItem]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT id, name, class_id FROM students WHERE class_id=? ORDER BY name",
                (class_id,),
            ).fetchall()
        return [StudentItem(id=row["id"], name=row["name"], class_id=row["class_id"]) for row in rows]

    def get_exams(self, class_id: int) -> List[ExamItem]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT id, name, class_id FROM exams WHERE class_id=? ORDER BY created_at",
                (class_id,),
            ).fetchall()
        return [ExamItem(id=row["id"], name=row["name"], class_id=row["class_id"]) for row in rows]

    def upsert_student(self, name: str, class_id: int) -> StudentItem:
        with self._connect() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO students(name, class_id) VALUES(?, ?)",
                (name, class_id),
            )
            row = conn.execute(
                "SELECT id, name, class_id FROM students WHERE name=? AND class_id=?",
                (name, class_id),
            ).fetchone()
        return StudentItem(id=row["id"], name=row["name"], class_id=row["class_id"])

    def upsert_exam(self, name: str, class_id: int) -> ExamItem:
        with self._connect() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO exams(name, class_id) VALUES(?, ?)",
                (name, class_id),
            )
            row = conn.execute(
                "SELECT id, name, class_id FROM exams WHERE name=? AND class_id=?",
                (name, class_id),
            ).fetchone()
        return ExamItem(id=row["id"], name=row["name"], class_id=row["class_id"])

    def insert_scores(self, records: Iterable[Dict[str, Any]]) -> None:
        with self._connect() as conn:
            conn.executemany(
                """
                INSERT INTO scores(
                    student_id, exam_id, subject, score, score_raw, rank,
                    total_score, total_raw, grade_rank, class_rank
                ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        r["student_id"],
                        r["exam_id"],
                        r["subject"],
                        r.get("score"),
                        r.get("score_raw"),
                        r.get("rank"),
                        r.get("total_score"),
                        r.get("total_raw"),
                        r.get("grade_rank"),
                        r.get("class_rank"),
                    )
                    for r in records
                ],
            )

    def get_scores_by_student(self, student_id: int) -> List[sqlite3.Row]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT s.*, e.name AS exam_name
                FROM scores s
                JOIN exams e ON s.exam_id = e.id
                WHERE s.student_id=?
                ORDER BY e.created_at
                """,
                (student_id,),
            ).fetchall()
        return rows

    def get_scores_for_students(self, student_ids: List[int]) -> List[sqlite3.Row]:
        if not student_ids:
            return []
        placeholders = ",".join(["?"] * len(student_ids))
        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT s.*, e.name AS exam_name, st.name AS student_name, st.class_id
                FROM scores s
                JOIN exams e ON s.exam_id = e.id
                JOIN students st ON s.student_id = st.id
                WHERE s.student_id IN ({placeholders})
                ORDER BY e.created_at
                """,
                student_ids,
            ).fetchall()
        return rows

    def delete_exam(self, exam_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM exams WHERE id=?", (exam_id,))

    def delete_student(self, student_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM students WHERE id=?", (student_id,))

    def rename_class(self, class_id: int, new_name: str) -> None:
        with self._connect() as conn:
            conn.execute("UPDATE classes SET name=? WHERE id=?", (new_name, class_id))

    def delete_class(self, class_id: int) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM classes WHERE id=?", (class_id,))

    def clear_class_data(self, class_id: int) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM scores WHERE student_id IN (SELECT id FROM students WHERE class_id=?)",
                (class_id,),
            )
            conn.execute("DELETE FROM exams WHERE class_id=?", (class_id,))
            conn.execute("DELETE FROM students WHERE class_id=?", (class_id,))

    def update_score_item(
        self,
        score_id: int,
        score: Optional[float],
        score_raw: Optional[float],
        rank: Optional[int],
        total_score: Optional[float],
        total_raw: Optional[float],
        grade_rank: Optional[int],
        class_rank: Optional[int],
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE scores
                SET score=?, score_raw=?, rank=?, total_score=?, total_raw=?, grade_rank=?, class_rank=?
                WHERE id=?
                """,
                (score, score_raw, rank, total_score, total_raw, grade_rank, class_rank, score_id),
            )

    def set_setting(self, key: str, value: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO settings(key, value) VALUES(?, ?)
                ON CONFLICT(key) DO UPDATE SET value=excluded.value
                """,
                (key, value),
            )

    def get_setting(self, key: str) -> Optional[str]:
        with self._connect() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else None

    def list_subjects(self, class_id: int) -> List[str]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT DISTINCT subject
                FROM scores s
                JOIN students st ON s.student_id = st.id
                WHERE st.class_id=?
                ORDER BY subject
                """,
                (class_id,),
            ).fetchall()
        return [row["subject"] for row in rows]

    def export_database(self, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with self._connect() as conn:
            backup = sqlite3.connect(output_path)
            conn.backup(backup)
            backup.close()

    def import_database(self, input_path: Path) -> None:
        input_path = Path(input_path)
        if not input_path.exists():
            raise FileNotFoundError(str(input_path))
        input_conn = sqlite3.connect(input_path)
        with self._connect() as conn:
            input_conn.backup(conn)
        input_conn.close()

    def get_all_scores(self) -> List[sqlite3.Row]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT s.*, e.name AS exam_name, st.name AS student_name, st.class_id
                FROM scores s
                JOIN exams e ON s.exam_id = e.id
                JOIN students st ON s.student_id = st.id
                ORDER BY e.created_at
                """
            ).fetchall()
        return rows
