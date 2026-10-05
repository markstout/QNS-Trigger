# SQL Database Schema: Tasks, Lists, and Calendar

This document outlines a relational SQL database schema for **CapIt / GrabIt**, representing the data structures currently managed via Google Drive spreadsheets and JSON files (`TasksAndLists` and `Calendar Events`).

The schema is designed to be compatible with standard relational databases such as **SQLite**, **PostgreSQL**, and **MySQL/MariaDB**.

---

## 1. Entity-Relationship Summary

```text
+-------------------+        1 : N        +-------------------+
|    categories     |-------------------->|       tasks       |
+-------------------+                     +-------------------+
| id (PK)           |                     | id (PK)           |
| account           |                     | account           |
| name              |                     | category_id (FK)  |
| is_permanent      |                     | text              |
| lock_completed    |                     | ...               |
+-------------------+                     +-------------------+

+-------------------+                     +-------------------+
|  calendar_events  |                     |   calendar_tags   |
+-------------------+                     +-------------------+
| id (PK)           |                     | name (PK)         |
| account           |                     | account           |
| title             |                     | color             |
| tag (FK)          |----->               +-------------------+
| ...               |
+-------------------+

+-------------------+ (Optional Normalized Table)
|   repeat_rules    |
+-------------------+
| id (PK)           |
| account           |
| entity_id (FK)    |
| entity_type       |
| frequency         |
| interval          |
| ends              |
| end_date          |
| end_occurrences   |
| current_occ       |
+-------------------+
```

---

## 2. Table Specifications

### 2.1 `categories` (Task Lists / Categories)
Stores task categories or list containers (e.g., "Today's Tasks", "Pending", "Groceries").

| Column | Type | Constraints | Description |
| :--- | :--- | :--- | :--- |
| `id` | VARCHAR(64) | PRIMARY KEY | Unique ID (e.g. `today`, `pending`, `cat_1716854930123`) |
| `account` | VARCHAR(255) | NOT NULL DEFAULT 'markstout@gmail.com' | Google Account email owning the record |
| `name` | VARCHAR(255) | NOT NULL | Display name of the category/list |
| `is_permanent` | BOOLEAN | DEFAULT FALSE | Protects system lists ("Today's Tasks", "Pending") from deletion |
| `lock_completed` | BOOLEAN | DEFAULT FALSE | Prevents archiving/clearing completed items for this list |
| `display_order` | INT | DEFAULT 0 | Ordering index for custom sorting |
| `created_at` | DATETIME | DEFAULT CURRENT_TIMESTAMP | Timestamp when category was created |

---

### 2.2 `tasks` (Task / Checklist Items)
Stores individual tasks associated with a category.

| Column | Type | Constraints | Description |
| :--- | :--- | :--- | :--- |
| `id` | VARCHAR(64) | PRIMARY KEY | UUID or stable hash ID |
| `account` | VARCHAR(255) | NOT NULL DEFAULT 'markstout@gmail.com' | Google Account email owning the record |
| `category_id` | VARCHAR(64) | FK -> `categories(id)` | Parent category ID |
| `text` | TEXT | NOT NULL | Main task description |
| `note` | TEXT | NULL | Detailed notes or body text |
| `url` | TEXT | NULL | Linked URL |
| `status` | VARCHAR(20) | DEFAULT 'open' | Lifecycle status (`open`, `completed`, `cancelled`) |
| `date_created` | DATETIME | DEFAULT CURRENT_TIMESTAMP | Creation timestamp |
| `date_completed` | DATETIME | NULL | Completion timestamp |
| `date_pending` | DATETIME | NULL | Timestamp when marked as pending |
| `date_cancelled` | DATETIME | NULL | Timestamp when cancelled |
| `is_cancelled` | BOOLEAN | DEFAULT FALSE | Cancellation flag |
| `completion_note` | TEXT | NULL | User note entered upon completion |
| `alarm_datetime` | DATETIME | NULL | Alarm date and time (`YYYY-MM-DD HH:MM`) |
| `alarm_type` | VARCHAR(20) | DEFAULT 'global' | Alert mode (`audio`, `notification`, `both`, `global`) |
| `associated_note_id` | VARCHAR(255) | NULL | Linked Google Document ID |
| `repeat_json` | TEXT | NULL | JSON string representing recurrence rules (or link to `repeat_rules`) |

---

### 2.3 `calendar_events` (Calendar Scheduling)
Stores calendar events with start/end times, tags, and recurrence.

| Column | Type | Constraints | Description |
| :--- | :--- | :--- | :--- |
| `id` | VARCHAR(64) | PRIMARY KEY | Event ID or stable hash ID |
| `account` | VARCHAR(255) | NOT NULL DEFAULT 'markstout@gmail.com' | Google Account email owning the record |
| `title` | VARCHAR(255) | NOT NULL | Event title/summary |
| `status` | VARCHAR(20) | DEFAULT 'open' | Status (`open`, `completed`, `cancelled`) |
| `event_date` | DATE | NOT NULL | Start date (`YYYY-MM-DD`) |
| `event_time` | TIME | NOT NULL | Start time (`HH:MM`) |
| `end_date` | DATE | NULL | End date (`YYYY-MM-DD`) |
| `end_time` | TIME | NULL | End time (`HH:MM`) |
| `location` | TEXT | NULL | Location / address string |
| `note` | TEXT | NULL | Brief note / description |
| `note_link` | TEXT | NULL | URL or document link |
| `tag` | VARCHAR(100) | NULL | Classification / Tag |
| `alarm_datetime` | DATETIME | NULL | Scheduled reminder timestamp |
| `alarm_type` | VARCHAR(20) | DEFAULT 'global' | Alert mode (`audio`, `notification`, `both`, `global`) |
| `associated_note_id` | VARCHAR(255) | NULL | Linked Google Document ID |
| `created_at` | DATETIME | DEFAULT CURRENT_TIMESTAMP | Creation timestamp |
| `completed_at` | DATETIME | NULL | Completion timestamp |
| `cancelled_at` | DATETIME | NULL | Cancellation timestamp |
| `completion_note` | TEXT | NULL | Note added on completion |
| `repeat_json` | TEXT | NULL | JSON string representing recurrence rules |

---

### 2.4 `repeat_rules` (Normalized Recurrence Rules)
Optional normalized table to store repeat configurations instead of embedding JSON.

| Column | Type | Constraints | Description |
| :--- | :--- | :--- | :--- |
| `id` | INTEGER | PRIMARY KEY AUTOINCREMENT | Surrogate key |
| `account` | VARCHAR(255) | NOT NULL DEFAULT 'markstout@gmail.com' | Google Account email owning the record |
| `entity_type` | VARCHAR(20) | NOT NULL | `task` or `calendar` |
| `entity_id` | VARCHAR(64) | NOT NULL | FK to `tasks(id)` or `calendar_events(id)` |
| `frequency` | VARCHAR(20) | NOT NULL | `Daily`, `Weekly`, `Monthly`, `Yearly` |
| `interval` | INT | DEFAULT 1 | Frequency multiplier (every N days/weeks) |
| `ends` | VARCHAR(20) | DEFAULT 'Never' | Termination type (`Never`, `On Date`, `After`) |
| `end_date` | DATE | NULL | Expiration date if `ends` = 'On Date' |
| `end_occurrences` | INT | NULL | Max occurrences if `ends` = 'After' |
| `current_occurrence` | INT | DEFAULT 0 | Count of completed occurrences |

---

### 2.5 `calendar_tags` (Tag Definitions)
Stores user-defined calendar event tag labels.

| Column | Type | Constraints | Description |
| :--- | :--- | :--- | :--- |
| `name` | VARCHAR(100) | PRIMARY KEY | Tag label (e.g., "Work", "Personal") |
| `account` | VARCHAR(255) | NOT NULL DEFAULT 'markstout@gmail.com' | Google Account email owning the record |
| `color` | VARCHAR(20) | NULL | Optional color hex code |

---

## 3. SQL Data Definition Language (DDL)

```sql
-- Disable foreign key checks during creation (SQLite syntax)
PRAGMA foreign_keys = ON;

-- -----------------------------------------------------
-- Table: categories
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS categories (
    id VARCHAR(64) PRIMARY KEY,
    account VARCHAR(255) NOT NULL DEFAULT 'markstout@gmail.com',
    name VARCHAR(255) NOT NULL,
    is_permanent BOOLEAN NOT NULL DEFAULT 0,
    lock_completed BOOLEAN NOT NULL DEFAULT 0,
    display_order INT DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_categories_account ON categories(account);

-- Seed Default Permanent Categories
INSERT OR IGNORE INTO categories (id, account, name, is_permanent, display_order) 
VALUES ('today', 'markstout@gmail.com', 'Today''s Tasks', 1, 0);

INSERT OR IGNORE INTO categories (id, account, name, is_permanent, display_order) 
VALUES ('pending', 'markstout@gmail.com', 'Pending', 1, 1);


-- -----------------------------------------------------
-- Table: tasks
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS tasks (
    id VARCHAR(64) PRIMARY KEY,
    account VARCHAR(255) NOT NULL DEFAULT 'markstout@gmail.com',
    category_id VARCHAR(64) NOT NULL,
    text TEXT NOT NULL,
    note TEXT,
    url TEXT,
    status VARCHAR(20) NOT NULL DEFAULT 'open',
    date_created DATETIME DEFAULT CURRENT_TIMESTAMP,
    date_completed DATETIME,
    date_pending DATETIME,
    date_cancelled DATETIME,
    is_cancelled BOOLEAN NOT NULL DEFAULT 0,
    completion_note TEXT,
    alarm_datetime DATETIME,
    alarm_type VARCHAR(20) DEFAULT 'global',
    associated_note_id VARCHAR(255),
    repeat_json TEXT,
    FOREIGN KEY (category_id) REFERENCES categories(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_tasks_account ON tasks(account);
CREATE INDEX IF NOT EXISTS idx_tasks_category ON tasks(category_id);
CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status);
CREATE INDEX IF NOT EXISTS idx_tasks_alarm ON tasks(alarm_datetime);


-- -----------------------------------------------------
-- Table: calendar_tags
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS calendar_tags (
    name VARCHAR(100) PRIMARY KEY,
    account VARCHAR(255) NOT NULL DEFAULT 'markstout@gmail.com',
    color VARCHAR(20)
);

CREATE INDEX IF NOT EXISTS idx_calendar_tags_account ON calendar_tags(account);


-- -----------------------------------------------------
-- Table: calendar_events
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS calendar_events (
    id VARCHAR(64) PRIMARY KEY,
    account VARCHAR(255) NOT NULL DEFAULT 'markstout@gmail.com',
    title VARCHAR(255) NOT NULL,
    status VARCHAR(20) NOT NULL DEFAULT 'open',
    event_date DATE NOT NULL,
    event_time TIME NOT NULL,
    end_date DATE,
    end_time TIME,
    location TEXT,
    note TEXT,
    note_link TEXT,
    tag VARCHAR(100),
    alarm_datetime DATETIME,
    alarm_type VARCHAR(20) DEFAULT 'global',
    associated_note_id VARCHAR(255),
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    completed_at DATETIME,
    cancelled_at DATETIME,
    completion_note TEXT,
    repeat_json TEXT,
    FOREIGN KEY (tag) REFERENCES calendar_tags(name) ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_calendar_events_account ON calendar_events(account);
CREATE INDEX IF NOT EXISTS idx_events_date ON calendar_events(event_date, event_time);
CREATE INDEX IF NOT EXISTS idx_events_status ON calendar_events(status);
CREATE INDEX IF NOT EXISTS idx_events_alarm ON calendar_events(alarm_datetime);


-- -----------------------------------------------------
-- Table: repeat_rules (Normalized Recurrence)
-- -----------------------------------------------------
CREATE TABLE IF NOT EXISTS repeat_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    account VARCHAR(255) NOT NULL DEFAULT 'markstout@gmail.com',
    entity_type VARCHAR(20) NOT NULL CHECK (entity_type IN ('task', 'calendar')),
    entity_id VARCHAR(64) NOT NULL,
    frequency VARCHAR(20) NOT NULL CHECK (frequency IN ('Daily', 'Weekly', 'Monthly', 'Yearly')),
    interval INT DEFAULT 1,
    ends VARCHAR(20) DEFAULT 'Never' CHECK (ends IN ('Never', 'On Date', 'After')),
    end_date DATE,
    end_occurrences INT,
    current_occurrence INT DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_repeat_rules_account ON repeat_rules(account);
CREATE INDEX IF NOT EXISTS idx_repeat_entity ON repeat_rules(entity_type, entity_id);
```

---

## 4. Useful SQL Queries

### Fetch All Active Tasks for "Today's Tasks" (Scoped to Account)
```sql
SELECT * FROM tasks 
WHERE account = 'markstout@gmail.com'
  AND category_id = 'today' 
  AND status = 'open' 
  AND date_completed IS NULL 
ORDER BY date_created DESC;
```

### Fetch Upcoming Calendar Events (From Today Onward, Scoped to Account)
```sql
SELECT * FROM calendar_events 
WHERE account = 'markstout@gmail.com'
  AND status = 'open' 
  AND event_date >= DATE('now')
ORDER BY event_date ASC, event_time ASC;
```

### Fetch All Pending Items Across All Lists (Scoped to Account)
```sql
SELECT t.*, c.name AS source_category_name 
FROM tasks t
JOIN categories c ON t.category_id = c.id AND t.account = c.account
WHERE t.account = 'markstout@gmail.com'
  AND (t.date_pending IS NOT NULL OR t.category_id = 'pending')
ORDER BY t.date_pending DESC;
```

### Fetch Items with Upcoming Alarms (Tasks + Calendar Combined, Scoped to Account)
```sql
SELECT 'task' AS item_type, id, text AS title, alarm_datetime, alarm_type 
FROM tasks 
WHERE account = 'markstout@gmail.com'
  AND alarm_datetime >= DATETIME('now') 
  AND status = 'open'

UNION ALL

SELECT 'calendar' AS item_type, id, title, alarm_datetime, alarm_type 
FROM calendar_events 
WHERE account = 'markstout@gmail.com'
  AND alarm_datetime >= DATETIME('now') 
  AND status = 'open'

ORDER BY alarm_datetime ASC;
```
