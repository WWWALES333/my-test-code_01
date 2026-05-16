from __future__ import annotations

from src.analysis_v14.schema import (
    DECISION_CONFIRMED,
    DECISION_PENDING_HUMAN,
    DECISION_UNCERTAIN,
    DECISION_VALUES,
    stable_hash,
)

TASK_STATUS_OPEN = "open"
TASK_STATUS_IN_REVIEW = "in_review"
TASK_STATUS_REVIEWED = "reviewed"

TASK_STATUS_VALUES = {
    TASK_STATUS_OPEN,
    TASK_STATUS_IN_REVIEW,
    TASK_STATUS_REVIEWED,
}

OWNER_TYPE_PERSON = "person"
OWNER_TYPE_GROUP = "group"
OWNER_TYPE_UNKNOWN = "unknown"

LEARNING_ERROR_REASON_VALUES = (
    "label_gap",
    "actor_boundary",
    "business_line_boundary",
    "ai_scope_boundary",
    "low_signal_noise",
    "context_loss",
    "parser_or_segmentation_error",
    "model_misread",
    "rule_threshold_issue",
    "other",
)

REVIEW_NECESSITY_VALUES = (
    "should_review",
    "could_auto_confirm",
    "could_auto_reject",
    "low_value_noise",
)

ACTIONABILITY_VALUES = (
    "actionable",
    "observe",
    "no_action",
)

ACTION_BUCKET_VALUES = (
    "product_pool",
    "sales_enablement_pool",
    "watchlist",
    "none",
)
