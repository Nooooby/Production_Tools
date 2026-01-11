# Unified Input Area Guide (Excel)

**Audience**: Supervisors preparing data inputs for scheduling and staffing.

This guide defines a **single, unified input area** in Excel using four tables:
`Departments`, `StaffPool`, `Stations`, and `BreakRules`. These tables should be
created as **Excel Tables** (Insert → Table) with the exact names below.

---

## 1) Table: `Departments`

**Purpose**: Define each department and its default shift window.

**Required fields**:
- `DeptID` (unique, required)
- `DeptName` (required)
- `ShiftStart` (required)
- `ShiftEnd` (required)

**Field definitions & format**:
- `DeptID`: Short text ID (e.g., TP, CU, BG). **Must be unique**.
- `DeptName`: Full department name (e.g., Tray Pack).
- `ShiftStart`: Time in `HH:MM` (24-hour) format (e.g., 06:30).
- `ShiftEnd`: Time in `HH:MM` (24-hour) format (e.g., 15:00).

**Validation rules**:
- No blanks in required fields.
- `DeptID` must be unique (no duplicates).
- `ShiftEnd` must be later than `ShiftStart`.

---

## 2) Table: `StaffPool`

**Purpose**: Define available staff, their skills, and rotation eligibility.

**Required fields**:
- `DeptID` (required, must match `Departments[DeptID]`)
- `Name` (required)
- `SkillType` (required)
- `Priority` (required)
- `CanRotate` (required)

**Field definitions & format**:
- `DeptID`: Must match a value in `Departments[DeptID]`.
- `Name`: Full name (text).
- `SkillType`: Skill category (e.g., Packer, Cutter, QA, Lead).
- `Priority`: Integer (e.g., 1–5, where 1 = highest priority).
- `CanRotate`: TRUE/FALSE (or Yes/No) to indicate rotation eligibility.

**Validation rules**:
- No blanks in required fields.
- `DeptID` must exist in `Departments`.
- `Priority` must be a positive integer within your chosen range (e.g., 1–5).
- `CanRotate` must be TRUE/FALSE (or Yes/No), not free text.

---

## 3) Table: `Stations`

**Purpose**: Define workstations and required staffing by skill.

**Required fields**:
- `DeptID` (required, must match `Departments[DeptID]`)
- `StationID` (required)
- `SkillRequired` (required)
- `Headcount` (required)

**Field definitions & format**:
- `DeptID`: Must match a value in `Departments[DeptID]`.
- `StationID`: Station code (text). **Unique within each DeptID**.
- `SkillRequired`: Skill category required at the station.
- `Headcount`: Positive integer (e.g., 1, 2, 3).

**Validation rules**:
- No blanks in required fields.
- `DeptID` must exist in `Departments`.
- `StationID` must be unique within the same `DeptID`.
- `Headcount` must be a positive integer (> 0).

---

## 4) Table: `BreakRules`

**Purpose**: Define break scheduling rules by department.

**Required fields**:
- `DeptID` (required, must match `Departments[DeptID]`)
- `BreakWindow` (required)
- `BatchCount` (required)
- `BatchSize` (required)
- `MaxConsecutive` (required)

**Field definitions & format**:
- `DeptID`: Must match a value in `Departments[DeptID]`.
- `BreakWindow`: Time range in `HH:MM-HH:MM` (24-hour) format.
- `BatchCount`: Integer ≥ 1 (number of break batches).
- `BatchSize`: Integer ≥ 1 (people per batch).
- `MaxConsecutive`: Integer ≥ 1 (max consecutive staff in a row without break).

**Validation rules**:
- No blanks in required fields.
- `DeptID` must exist in `Departments`.
- `BreakWindow` must be a valid time range and the end time must be later than the start time.
- `BatchCount`, `BatchSize`, `MaxConsecutive` must be positive integers.

---

# Example Filling (Supervisor Reference)

> **Note**: The following examples are for reference only. Do not treat them as script logic.

## Example: `Departments`

| DeptID | DeptName   | ShiftStart | ShiftEnd |
|------:|------------|-----------:|---------:|
| TP    | Tray Pack  | 06:30      | 15:00    |
| CU    | Cut-Up     | 06:00      | 14:30    |
| BG    | Bagging    | 07:00      | 15:30    |

## Example: `StaffPool`

| DeptID | Name         | SkillType | Priority | CanRotate |
|------:|--------------|-----------|---------:|----------:|
| TP    | Alex Chen    | Packer    | 1        | TRUE      |
| TP    | Maria Lopez  | QA        | 2        | FALSE     |
| CU    | David Kim    | Cutter    | 1        | TRUE      |
| BG    | Priya Singh  | Packer    | 3        | TRUE      |

## Example: `Stations`

| DeptID | StationID | SkillRequired | Headcount |
|------:|-----------|---------------|----------:|
| TP    | TP-01     | Packer        | 2         |
| TP    | TP-02     | QA            | 1         |
| CU    | CU-01     | Cutter        | 3         |
| BG    | BG-01     | Packer        | 2         |

## Example: `BreakRules`

| DeptID | BreakWindow | BatchCount | BatchSize | MaxConsecutive |
|------:|-------------|-----------:|----------:|---------------:|
| TP    | 09:30-11:00  | 3          | 4         | 6              |
| CU    | 09:00-10:30  | 2          | 5         | 5              |
| BG    | 10:00-11:30  | 2          | 3         | 4              |
