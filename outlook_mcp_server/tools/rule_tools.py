"""Rule management MCP tools for Outlook."""

import json
from typing import Any, Dict, List, Optional, Union
from ..backend.outlook_session import OutlookSessionManager
from ..backend.outlook_session.folder_operations import FolderOperations
from ..backend.outlook_session.rule_operations import RuleOperations


def _parse_list(value: Any) -> Optional[List[str]]:
    """Coerce MCP input to List[str] regardless of how FastMCP serialized it.

    FastMCP may transmit Optional[List[str]] parameters as:
      - None          (not provided)
      - list          (correct — pass through)
      - str           (JSON-encoded list e.g. '["a", "b"]' or single value "a")

    This helper normalizes all three cases so rule_operations always receives
    a proper Python list or None.
    """
    if value is None:
        return None
    if isinstance(value, list):
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.startswith("["):
            try:
                result = json.loads(stripped)
                if isinstance(result, list):
                    return [str(item) for item in result]
            except (json.JSONDecodeError, ValueError):
                pass
        # Single bare string — treat as one-item list
        return [stripped] if stripped else None
    return None


def list_rules_tool() -> Dict[str, Any]:
    """List all Outlook rules with their conditions, actions, and enabled state.

    Returns:
        dict: Formatted list of all rules showing name, execution order,
              enabled status, conditions (from, subject, sent_to), and
              actions (move_to_folder, assign_category, delete, stop).

    Use this tool to:
      - Audit existing rules before creating new ones
      - Verify rule paths after folder restructure
      - Identify disabled or broken rules
    """
    try:
        with OutlookSessionManager() as sm:
            folder_ops = FolderOperations(sm)
            rule_ops = RuleOperations(sm, folder_ops)
            rules = rule_ops.list_rules()

            if not rules:
                return {"type": "text", "text": "No rules found."}

            lines = [f"Found {len(rules)} rules:\n"]
            for r in rules:
                status = "✅ ON" if r.get("enabled") else "❌ OFF"
                lines.append(f"[{r.get('execution_order', '?')}] {r['name']} — {status}")

                conds = r.get("conditions", {})
                if conds.get("from_addresses"):
                    lines.append(f"    FROM: {', '.join(conds['from_addresses'])}")
                if conds.get("subject_contains"):
                    lines.append(f"    SUBJECT CONTAINS: {', '.join(conds['subject_contains'])}")
                if conds.get("sent_to"):
                    lines.append(f"    SENT TO: {', '.join(conds['sent_to'])}")
                if conds.get("only_to_me"):
                    lines.append(f"    ONLY TO ME: true")
                if conds.get("to_or_cc_me"):
                    lines.append(f"    TO OR CC ME: true")
                if conds.get("not_to_or_cc_me"):
                    lines.append(f"    NOT TO OR CC ME: true")

                acts = r.get("actions", {})
                if acts.get("move_to_folder"):
                    lines.append(f"    → MOVE TO: {acts['move_to_folder']}")
                if acts.get("assign_category"):
                    lines.append(f"    → CATEGORY: {acts['assign_category']}")
                if acts.get("delete"):
                    lines.append(f"    → DELETE")
                if acts.get("stop_processing"):
                    lines.append(f"    → STOP PROCESSING")
                lines.append("")

            return {"type": "text", "text": "\n".join(lines)}
    except Exception as e:
        return {"type": "text", "text": f"Error listing rules: {str(e)}"}


def create_rule_tool(
    name: str,
    from_addresses: Optional[str] = None,
    subject_contains: Optional[str] = None,
    sent_to_me: bool = False,
    to_or_cc_me: bool = False,
    not_to_or_cc_me: bool = False,
    move_to_folder: Optional[str] = None,
    assign_category: Optional[str] = None,
    delete_message: bool = False,
    stop_processing: bool = True,
    enabled: bool = True,
) -> Dict[str, Any]:
    """Create a new Outlook receive rule.

    Args:
        name: Unique rule name
        from_addresses: Sender addresses to match. Pass as JSON array string:
                        '["pietranm@couche-tard.com"]'
                        Supports partial domain matching: '["@eaton.com"]'
                        Single address also accepted: "pietranm@couche-tard.com"
        subject_contains: Subject text strings (any match triggers — OR logic).
                          Pass as JSON array string: '["P1", "P2", "critical"]'
                          Single value also accepted: "URGENT"
        sent_to_me: True = only trigger when Dave is the sole To: recipient
        to_or_cc_me: True = only trigger when Dave is in To: or CC:
        not_to_or_cc_me: True = only trigger when Dave is NOT in To: or CC:
                         Use this for the Universal Noise Filter.
        move_to_folder: Destination folder path e.g. "Inbox/MANAGEMENT/Mario"
                        Use "Inbox/FolderName" format — no email prefix.
        assign_category: Outlook category name e.g. "URGENT"
        delete_message: True = delete the message
        stop_processing: Stop evaluating further rules after this one (default True)
        enabled: Rule is active immediately (default True)

    Returns:
        dict: Success or error message

    Note:
        Must provide at least one condition AND at least one action.
        Rules.Save(True) is called automatically to push to Exchange server.
        NEVER create a rule named "[HEALTH]" — that is ETL infrastructure.
    """
    try:
        from_list = _parse_list(from_addresses)
        subject_list = _parse_list(subject_contains)

        with OutlookSessionManager() as sm:
            folder_ops = FolderOperations(sm)
            rule_ops = RuleOperations(sm, folder_ops)
            result = rule_ops.create_rule(
                name=name,
                from_addresses=from_list,
                subject_contains=subject_list,
                sent_to_me=sent_to_me,
                to_or_cc_me=to_or_cc_me,
                not_to_or_cc_me=not_to_or_cc_me,
                move_to_folder=move_to_folder,
                assign_category=assign_category,
                delete_message=delete_message,
                stop_processing=stop_processing,
                enabled=enabled,
            )
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error creating rule: {str(e)}"}


def update_rule_tool(
    rule_name: str,
    new_name: Optional[str] = None,
    from_addresses: Optional[str] = None,
    subject_contains: Optional[str] = None,
    move_to_folder: Optional[str] = None,
    assign_category: Optional[str] = None,
    enabled: Optional[bool] = None,
) -> Dict[str, Any]:
    """Update an existing Outlook rule.

    Only parameters you provide are changed — omitted parameters are untouched.

    Args:
        rule_name: Exact name of the rule to update (case-insensitive)
        new_name: Rename the rule
        from_addresses: Replace From addresses. Pass as JSON array string:
                        '["addr1@example.com", "addr2@example.com"]'
                        Single address also accepted: "addr@example.com"
        subject_contains: Replace Subject text list. Pass as JSON array string:
                          '["PDU verification", "PDU VALIDATION", "PDU Validation"]'
                          Single value also accepted: "PDU"
        move_to_folder: Update destination folder path
        assign_category: Update or add category assignment
        enabled: True to enable, False to disable

    Returns:
        dict: Success or error message
    """
    try:
        from_list = _parse_list(from_addresses)
        subject_list = _parse_list(subject_contains)

        with OutlookSessionManager() as sm:
            folder_ops = FolderOperations(sm)
            rule_ops = RuleOperations(sm, folder_ops)
            result = rule_ops.update_rule(
                rule_name=rule_name,
                new_name=new_name,
                from_addresses=from_list,
                subject_contains=subject_list,
                move_to_folder=move_to_folder,
                assign_category=assign_category,
                enabled=enabled,
            )
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error updating rule: {str(e)}"}


def delete_rule_tool(rule_name: str) -> Dict[str, Any]:
    """Delete an Outlook rule by name.

    Args:
        rule_name: Exact name of the rule to delete (case-insensitive)

    Returns:
        dict: Success or error message

    Warning:
        This is permanent. Use update_rule_tool with enabled=False to
        temporarily disable a rule instead.

    HARD PROTECTION: Rule '[HEALTH]' cannot be deleted — it is ETL infrastructure.
    """
    try:
        with OutlookSessionManager() as sm:
            folder_ops = FolderOperations(sm)
            rule_ops = RuleOperations(sm, folder_ops)
            result = rule_ops.delete_rule(rule_name)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error deleting rule: {str(e)}"}


def reorder_rule_tool(rule_name: str, new_position: int) -> Dict[str, Any]:
    """Move a rule to a specific execution position.

    Rules execute in order. Lower position = runs first.
    The Universal Noise Filter must always be the LAST rule (highest position).

    Args:
        rule_name: Exact name of the rule to reorder (case-insensitive)
        new_position: Target position (1-based). Use 999 to push a rule last.

    Returns:
        dict: Success or error message

    Example:
        reorder_rule_tool("Universal Noise Filter", 999)
        → Ensures noise filter runs after all specific rules
    """
    try:
        with OutlookSessionManager() as sm:
            folder_ops = FolderOperations(sm)
            rule_ops = RuleOperations(sm, folder_ops)
            result = rule_ops.reorder_rule(rule_name, new_position)
            return {"type": "text", "text": result}
    except Exception as e:
        return {"type": "text", "text": f"Error reordering rule: {str(e)}"}
