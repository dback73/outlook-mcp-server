"""
Outlook rule management via COM.

Supports creating, reading, updating, deleting, and reordering
Outlook receive rules using the Rules object model.

COM object model reference:
  namespace.DefaultStore.GetRules() -> Rules
  Rules.Create(name, olRuleReceive=1) -> Rule
  Rule.Conditions:
    .From           -> AddressRuleCondition  (.Address list, .Enabled)
    .Subject        -> TextRuleCondition     (.Text list, .Enabled)
    .SentTo         -> AddressRuleCondition  (.Address list, .Enabled)
    .OnlyToMe       -> BooleanRuleCondition  (.Enabled)
    .ToOrCc         -> BooleanRuleCondition  (.Enabled — "Dave in To or CC")
    .NotToOrCc      -> BooleanRuleCondition  (.Enabled — "Dave NOT in To or CC")
  Rule.Actions:
    .MoveToFolder   -> MoveOrCopyRuleAction  (.Folder, .Enabled)
    .AssignToCategory -> RuleAction          (.Categories str, .Enabled)
    .Delete         -> RuleAction            (.Enabled)
    .Stop           -> RuleAction            (.Enabled)
  Rules.Save(True)  -> commits to Exchange server
"""

import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

# COM constants
OL_RULE_RECEIVE = 1  # olRuleReceive


class RuleOperations:
    """Handles all Outlook rule management via COM."""

    def __init__(self, session_manager, folder_operations):
        self.session_manager = session_manager
        self.folder_ops = folder_operations

    def _get_rules(self):
        """Get the Rules collection from the default store."""
        return self.session_manager.outlook_namespace.DefaultStore.GetRules()

    def _find_rule(self, rules, rule_name: str):
        """Find a rule by name (case-insensitive)."""
        for i in range(1, rules.Count + 1):
            rule = rules.Item(i)
            if rule.Name.lower() == rule_name.lower():
                return rule
        return None

    def _extract_rule_info(self, rule) -> Dict[str, Any]:
        """Extract readable info from a rule COM object."""
        info = {
            "name": rule.Name,
            "enabled": rule.Enabled,
            "execution_order": rule.ExecutionOrder,
            "conditions": {},
            "actions": {},
        }

        # Conditions
        try:
            cond = rule.Conditions
            # From addresses
            try:
                if cond.From.Enabled:
                    addrs = []
                    for j in range(1, cond.From.Address.Count + 1):
                        try:
                            addrs.append(cond.From.Address.Item(j).Address)
                        except Exception:
                            addrs.append(str(cond.From.Address.Item(j)))
                    info["conditions"]["from_addresses"] = addrs
            except Exception:
                pass

            # Subject contains
            try:
                if cond.Subject.Enabled:
                    texts = list(cond.Subject.Text)
                    info["conditions"]["subject_contains"] = texts
            except Exception:
                pass

            # SentTo
            try:
                if cond.SentTo.Enabled:
                    addrs = []
                    for j in range(1, cond.SentTo.Address.Count + 1):
                        try:
                            addrs.append(cond.SentTo.Address.Item(j).Address)
                        except Exception:
                            addrs.append(str(cond.SentTo.Address.Item(j)))
                    info["conditions"]["sent_to"] = addrs
            except Exception:
                pass

            # OnlyToMe
            try:
                if cond.OnlyToMe.Enabled:
                    info["conditions"]["only_to_me"] = True
            except Exception:
                pass

            # ToOrCc (Dave in To or CC)
            try:
                if cond.ToOrCc.Enabled:
                    info["conditions"]["to_or_cc_me"] = True
            except Exception:
                pass

            # NotToOrCc (Dave NOT in To or CC)
            try:
                if cond.NotToOrCc.Enabled:
                    info["conditions"]["not_to_or_cc_me"] = True
            except Exception:
                pass

        except Exception as e:
            info["conditions"]["_error"] = str(e)

        # Actions
        try:
            acts = rule.Actions
            # MoveToFolder
            try:
                if acts.MoveToFolder.Enabled:
                    info["actions"]["move_to_folder"] = acts.MoveToFolder.Folder.FolderPath
            except Exception:
                pass

            # AssignToCategory
            try:
                if acts.AssignToCategory.Enabled:
                    info["actions"]["assign_category"] = acts.AssignToCategory.Categories
            except Exception:
                pass

            # Delete
            try:
                if acts.Delete.Enabled:
                    info["actions"]["delete"] = True
            except Exception:
                pass

            # Stop
            try:
                if acts.Stop.Enabled:
                    info["actions"]["stop_processing"] = True
            except Exception:
                pass

        except Exception as e:
            info["actions"]["_error"] = str(e)

        return info

    # ── Public Methods ──────────────────────────────────────────────────────

    def list_rules(self) -> List[Dict[str, Any]]:
        """Return all rules with their conditions and actions."""
        rules = self._get_rules()
        result = []
        for i in range(1, rules.Count + 1):
            try:
                rule = rules.Item(i)
                result.append(self._extract_rule_info(rule))
            except Exception as e:
                result.append({"name": f"Rule {i}", "_error": str(e)})
        return result

    def create_rule(
        self,
        name: str,
        from_addresses: Optional[List[str]] = None,
        subject_contains: Optional[List[str]] = None,
        sent_to_me: bool = False,
        to_or_cc_me: bool = False,
        not_to_or_cc_me: bool = False,
        move_to_folder: Optional[str] = None,
        assign_category: Optional[str] = None,
        delete_message: bool = False,
        stop_processing: bool = True,
        enabled: bool = True,
    ) -> str:
        """
        Create a new Outlook receive rule.

        Args:
            name: Rule name (must be unique)
            from_addresses: List of sender email addresses to match
            subject_contains: List of subject text strings (OR logic within list)
            sent_to_me: True = only when Dave is the sole To: recipient
            to_or_cc_me: True = only when Dave is in To or CC
            not_to_or_cc_me: True = only when Dave is NOT in To or CC
                             (used for Universal Noise Filter — catches bulk/CC'd mail)
            move_to_folder: Destination folder path e.g. "Inbox/MANAGEMENT/Mario"
            assign_category: Category name e.g. "URGENT"
            delete_message: True = delete the message
            stop_processing: True = stop processing more rules after this one (default True)
            enabled: True = rule is active (default True)

        Returns:
            Success or error message
        """
        rules = self._get_rules()

        # Check for duplicate name
        if self._find_rule(rules, name):
            return f"Error: Rule '{name}' already exists. Use update_rule_tool to modify it."

        # Validate — must have at least one condition and one action
        has_condition = any([from_addresses, subject_contains, sent_to_me, to_or_cc_me, not_to_or_cc_me])
        has_action = any([move_to_folder, assign_category, delete_message])
        if not has_condition:
            return "Error: Must specify at least one condition (from_addresses, subject_contains, sent_to_me, to_or_cc_me, or not_to_or_cc_me)"
        if not has_action:
            return "Error: Must specify at least one action (move_to_folder, assign_category, or delete_message)"

        try:
            rule = rules.Create(name, OL_RULE_RECEIVE)

            # ── Conditions ──
            if from_addresses:
                from_cond = rule.Conditions.From
                from_cond.Enabled = True
                for addr in from_addresses:
                    recip = self.session_manager.outlook_namespace.CreateRecipient(addr)
                    recip.Resolve()
                    from_cond.Address.Add(recip)

            if subject_contains:
                subj_cond = rule.Conditions.Subject
                subj_cond.Enabled = True
                subj_cond.Text = subject_contains

            if sent_to_me:
                rule.Conditions.OnlyToMe.Enabled = True

            if to_or_cc_me:
                rule.Conditions.ToOrCc.Enabled = True

            if not_to_or_cc_me:
                rule.Conditions.NotToOrCc.Enabled = True

            # ── Actions ──
            if move_to_folder:
                target = self.folder_ops.get_folder(move_to_folder)
                if not target:
                    return f"Error: Destination folder '{move_to_folder}' not found"
                move_action = rule.Actions.MoveToFolder
                move_action.Enabled = True
                move_action.Folder = target

            if assign_category:
                cat_action = rule.Actions.AssignToCategory
                cat_action.Enabled = True
                cat_action.Categories = assign_category

            if delete_message:
                rule.Actions.Delete.Enabled = True

            if stop_processing:
                rule.Actions.Stop.Enabled = True

            rule.Enabled = enabled

            # Save to Exchange server
            rules.Save(True)

            logger.info(f"Created rule '{name}'")
            return f"Rule '{name}' created successfully"

        except Exception as e:
            logger.error(f"Error creating rule '{name}': {e}")
            return f"Error creating rule '{name}': {str(e)}"

    def update_rule(
        self,
        rule_name: str,
        new_name: Optional[str] = None,
        from_addresses: Optional[List[str]] = None,
        subject_contains: Optional[List[str]] = None,
        move_to_folder: Optional[str] = None,
        assign_category: Optional[str] = None,
        enabled: Optional[bool] = None,
    ) -> str:
        """
        Update an existing rule.

        Only provided parameters are updated — omitted parameters are unchanged.

        Args:
            rule_name: Name of the rule to update
            new_name: Rename the rule
            from_addresses: Replace From condition addresses (replaces, not appends)
            subject_contains: Replace Subject condition texts (replaces, not appends)
            move_to_folder: Update destination folder path
            assign_category: Update or add category assignment
            enabled: Enable or disable the rule

        Returns:
            Success or error message
        """
        rules = self._get_rules()
        rule = self._find_rule(rules, rule_name)
        if not rule:
            return f"Error: Rule '{rule_name}' not found"

        try:
            if new_name is not None:
                rule.Name = new_name

            if enabled is not None:
                rule.Enabled = enabled

            if from_addresses is not None:
                from_cond = rule.Conditions.From
                from_cond.Enabled = True
                # Clear existing addresses
                while from_cond.Address.Count > 0:
                    from_cond.Address.Remove(1)
                # Add new addresses
                for addr in from_addresses:
                    recip = self.session_manager.outlook_namespace.CreateRecipient(addr)
                    recip.Resolve()
                    from_cond.Address.Add(recip)

            if subject_contains is not None:
                subj_cond = rule.Conditions.Subject
                subj_cond.Enabled = True
                subj_cond.Text = subject_contains

            if move_to_folder is not None:
                target = self.folder_ops.get_folder(move_to_folder)
                if not target:
                    return f"Error: Destination folder '{move_to_folder}' not found"
                move_action = rule.Actions.MoveToFolder
                move_action.Enabled = True
                move_action.Folder = target

            if assign_category is not None:
                cat_action = rule.Actions.AssignToCategory
                cat_action.Enabled = True
                cat_action.Categories = assign_category

            rules.Save(True)
            display_name = new_name if new_name else rule_name
            logger.info(f"Updated rule '{rule_name}'")
            return f"Rule '{display_name}' updated successfully"

        except Exception as e:
            logger.error(f"Error updating rule '{rule_name}': {e}")
            return f"Error updating rule '{rule_name}': {str(e)}"

    def delete_rule(self, rule_name: str) -> str:
        """
        Delete a rule by name.

        Args:
            rule_name: Exact name of the rule to delete

        Returns:
            Success or error message
        """
        # Hard protection — never delete the ETL health rule
        if "[HEALTH]" in rule_name:
            return "Error: Rule '[HEALTH]' is ETL infrastructure and cannot be deleted via MCP."

        rules = self._get_rules()
        rule_index = None
        for i in range(1, rules.Count + 1):
            if rules.Item(i).Name.lower() == rule_name.lower():
                rule_index = i
                break
        if rule_index is None:
            return f"Error: Rule '{rule_name}' not found"

        try:
            rules.Remove(rule_index)
            rules.Save(True)
            logger.info(f"Deleted rule '{rule_name}'")
            return f"Rule '{rule_name}' deleted successfully"
        except Exception as e:
            logger.error(f"Error deleting rule '{rule_name}': {e}")
            return f"Error deleting rule '{rule_name}': {str(e)}"

    def reorder_rule(self, rule_name: str, new_position: int) -> str:
        """
        Move a rule to a specific execution position.

        Position 1 = runs first. Use a high number (e.g. 999) to push a rule last.
        The Universal Noise Filter must always be last.

        Args:
            rule_name: Name of the rule to reorder
            new_position: Target position (1-based)

        Returns:
            Success or error message
        """
        rules = self._get_rules()
        rule = self._find_rule(rules, rule_name)
        if not rule:
            return f"Error: Rule '{rule_name}' not found"

        try:
            rule.ExecutionOrder = new_position
            rules.Save(True)
            logger.info(f"Moved rule '{rule_name}' to position {new_position}")
            return f"Rule '{rule_name}' moved to execution position {new_position}"
        except Exception as e:
            logger.error(f"Error reordering rule '{rule_name}': {e}")
            return f"Error reordering rule '{rule_name}': {str(e)}"
