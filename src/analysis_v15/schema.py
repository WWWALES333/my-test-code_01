from __future__ import annotations

from src.analysis_v14.schema import (
    DECISION_CONFIRMED,
    DECISION_PENDING_HUMAN,
    DECISION_UNCERTAIN,
    DECISION_VALUES,
    stable_hash,
)

TASK_STATUS_OPEN = "open"
TASK_STATUS_REVIEWED = "reviewed"

TASK_STATUS_VALUES = {
    TASK_STATUS_OPEN,
    TASK_STATUS_REVIEWED,
}

OWNER_TYPE_PERSON = "person"
OWNER_TYPE_GROUP = "group"
OWNER_TYPE_UNKNOWN = "unknown"

