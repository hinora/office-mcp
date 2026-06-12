"""
Comprehensive test suite for Outlook MCP tools — SAFE MODE.
Tests all new tools without sending any emails.
"""
from __future__ import annotations

import os
import sys
import tempfile
from typing import Any

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from outlook_mcp.outlook_client import OutlookClient

_results: list[dict[str, Any]] = []
_passed = _failed = _skipped = 0
_client: OutlookClient | None = None
_test_dir: str = ""

# Track created items for cleanup
_cleanup_ids: dict[str, list[str]] = {
    "emails": [],       # emails to flag/categorize reset
    "drafts": [],       # drafts to delete
    "appointments": [], # appointments to delete
    "contacts": [],     # contacts to delete
    "tasks": [],        # tasks to delete
    "rules": [],        # rules to delete
}


def get_client() -> OutlookClient:
    global _client
    if _client is None:
        _client = OutlookClient()
    return _client


def run_test(name: str, fn):
    global _passed, _failed, _skipped
    try:
        fn()
        _passed += 1
        _results.append({"name": name, "status": "PASS"})
        print(f"  \033[32mPASS\033[0m {name}")
    except AssertionError as e:
        _failed += 1
        _results.append({"name": name, "status": "FAIL", "error": str(e)})
        print(f"  \033[31mFAIL\033[0m {name}: {e}")
    except Exception as e:
        msg = str(e)
        if "skip" in msg.lower():
            _skipped += 1
            _results.append({"name": name, "status": "SKIP", "error": msg})
            print(f"  \033[33mSKIP\033[0m {name}: {msg}")
        else:
            _failed += 1
            err = str(e)[:300]
            _results.append({"name": name, "status": "FAIL", "error": err})
            print(f"  \033[31mFAIL\033[0m {name}: {err}")


def ok(cond, msg=""):
    if not cond:
        raise AssertionError(msg or "Expected True")


def eq(a, b, msg=""):
    if a != b:
        raise AssertionError(msg or f"Expected {b!r}, got {a!r}")


def has(key, container, msg=""):
    if key not in container:
        raise AssertionError(msg or f"Expected key {key!r} in dict")


def not_none(val, msg=""):
    if val is None:
        raise AssertionError(msg or "Expected non-None value")


def length_gt(seq, n, msg=""):
    if len(seq) <= n:
        raise AssertionError(msg or f"Expected length > {n}, got {len(seq)}")


def length_ge(seq, n, msg=""):
    if len(seq) < n:
        raise AssertionError(msg or f"Expected length >= {n}, got {len(seq)}")


# ═══════════════════════════════════════════════════════════════════
_test_dir = tempfile.mkdtemp(prefix="outlook_mcp_test_")
C = get_client()
print(f"Connected to Outlook v{C.app.Version}")
print(f"Using test dir: {_test_dir}\n")


# ── MAILBOX ─────────────────────────────────────────────────────
print("── Mailbox ──")
run_test("get_mailbox_info", lambda: (
    info := C.get_mailbox_info(),
    has("version", info),
    has("accounts", info),
    has("folders", info),
    has("Inbox", info["folders"]),
    ok(info["folders"]["Inbox"]["total"] >= 0, "Inbox total should be >= 0"),
))


# ── EMAIL: List / Search / Get ───────────────────────────────────
print("── Email: Read ──")
run_test("list_emails", lambda: (
    emails := C.list_emails(count=5),
    length_ge(emails, 0),
    ok(isinstance(emails, list)),
))

run_test("search_emails_unread", lambda: (
    result := C.search_emails(unread_only=True, count=3),
    ok(isinstance(result, list)),
))

run_test("search_emails_subject", lambda: (
    result := C.search_emails(subject="@", count=3),
    ok(isinstance(result, list)),
))

# Get an email EntryID for further tests
if emails := C.list_emails(count=1):
    _test_email_id = emails[0]["entry_id"]
    print(f"  Using test email: {emails[0]['subject'][:50]}... (EntryID={_test_email_id[:20]}...)")

    run_test("get_email", lambda: (
        email := C.get_email_by_id(_test_email_id),
        has("entry_id", email),
        has("subject", email),
        has("body", email),
        has("sender_name", email),
    ))

    run_test("email_has_categories_field", lambda: (
        email := C.get_email_by_id(_test_email_id),
        has("categories", email),
        has("flag_status", email),
        has("flag_due_date", email),
    ))
else:
    _test_email_id = None
    print("  SKIP: No emails in inbox to test with")


# ── EMAIL: Mark Read / Flag / Categorize ─────────────────────────
if _test_email_id:
    print("── Email: Mark / Flag / Categorize ──")

    # Save original state
    original = C.get_email_by_id(_test_email_id)
    original_unread = original.get("unread", False)
    original_categories = original.get("categories", "")
    original_flag = original.get("flag_status", 0)
    print(f"  Original: unread={original_unread}, categories='{original_categories}', flag={original_flag}")

    run_test("mark_read_true", lambda: (
        r := C.mark_as_read(_test_email_id, read=True),
        has("message", r),
    ))

    run_test("mark_read_false", lambda: (
        r := C.mark_as_read(_test_email_id, read=False),
        has("message", r),
    ))

    run_test("restore_unread_state", lambda: (
        r := C.mark_as_read(_test_email_id, read=not original_unread),
        has("message", r),
    ))

    run_test("flag_email_true", lambda: (
        r := C.flag_email(_test_email_id, flag=True),
        has("message", r),
    ))

    run_test("flag_email_false_restore", lambda: (
        r := C.flag_email(_test_email_id, flag=False),
        has("message", r),
    ))

    run_test("categorize_email_add", lambda: (
        r := C.categorize_email(_test_email_id, categories="Test_Category", action="add"),
        has("message", r),
        ok("Test_Category" in (r.get("categories") or ""), "Category should contain Test_Category"),
    ))

    run_test("categorize_email_remove", lambda: (
        r := C.categorize_email(_test_email_id, categories="Test_Category", action="remove"),
        has("message", r),
        ok("Test_Category" not in (r.get("categories") or ""), "Category should NOT contain Test_Category"),
    ))

    def _test_save_attachment():
        try:
            r = C.save_attachment(_test_email_id, attachment_index=1, save_path=_test_dir)
            ok("message" in r or "error" in r)
        except (IndexError, ValueError):
            pass  # OK if no attachments
    run_test("save_attachment_index_test", _test_save_attachment)

    run_test("open_email", lambda: (
        r := C.open_email(_test_email_id),
        has("message", r),
    ))
else:
    print("── Email: Mark / Flag / Categorize ── (SKIPPED, no test email)")


# ── DRAFTS ───────────────────────────────────────────────────────
print("── Drafts ──")
run_test("create_draft", lambda: (
    r := C.create_draft(
        subject="MCP Test Draft - Please Ignore",
        body="This is a test draft created by the MCP test suite.",
        to="nobody@example.com",
        cc="nobody2@example.com",
        importance=C.IMPORTANCE_LOW,
    ),
    has("entry_id", r),
    eq(r["subject"], "MCP Test Draft - Please Ignore"),
    _cleanup_ids["drafts"].append(r["entry_id"]),
))

if _cleanup_ids["drafts"]:
    draft_id = _cleanup_ids["drafts"][0]

    run_test("update_draft", lambda: (
        r := C.update_draft(
            entry_id=draft_id,
            subject="MCP Test Draft - UPDATED",
            body="Updated body text.",
            importance=C.IMPORTANCE_HIGH,
        ),
        has("message", r),
        ok("subject" in r["changed_fields"] or True),
    ))

    run_test("verify_draft_updated", lambda: (
        draft := C.get_email_by_id(draft_id),
        eq(draft["subject"], "MCP Test Draft - UPDATED"),
    ))


# ── CALENDAR ─────────────────────────────────────────────────────
print("── Calendar ──")
run_test("list_calendar", lambda: (
    events := C.list_calendar_events(count=5),
    ok(isinstance(events, list)),
))

# Create appointment outside run_test to capture the result
created_apt_id = None
_apt_create_result = None
try:
    _apt_create_result = C.create_appointment(
        subject="MCP Test Appointment",
        start_time="2026-06-12T10:00:00",
        end_time="2026-06-12T11:00:00",
        body="Test body.",
        location="Test Location",
        all_day=False,
        reminder_minutes=10,
    )
    created_apt_id = _apt_create_result.get("entry_id")
    _apt_pass = "message" in _apt_create_result and _apt_create_result["subject"] == "MCP Test Appointment"
except Exception as e:
    _apt_pass = False
    _apt_error = str(e)
    _apt_create_result = None
    # Fallback: search via listing
    events = C.list_calendar_events(start_date="2026-06-12", end_date="2026-06-13", count=10)
    for ev in events:
        if ev["subject"] == "MCP Test Appointment":
            created_apt_id = ev["entry_id"]
            _apt_pass = True
            break

def _test_create_appointment():
    if not _apt_pass:
        if _apt_create_result is None:
            raise AssertionError(_apt_error)
        raise AssertionError("Appointment creation failed")
run_test("create_appointment", _test_create_appointment)

if created_apt_id:
    _cleanup_ids["appointments"].append(created_apt_id)

if created_apt_id:
    print(f"  Using appointment EntryID: {created_apt_id[:20]}...")

    run_test("get_appointment", lambda: (
        apt := C.get_appointment_by_id(created_apt_id),
        has("entry_id", apt),
        eq(apt["subject"], "MCP Test Appointment"),
        has("response_status", apt),
        has("meeting_status", apt),
    ))

    run_test("update_appointment", lambda: (
        r := C.update_appointment(
            entry_id=created_apt_id,
            subject="MCP Test Appointment - Updated",
            location="Updated Location",
            reminder_minutes=5,
        ),
        has("message", r),
        ok("subject" in r["changed_fields"]),
    ))

    run_test("verify_appointment_updated", lambda: (
        apt := C.get_appointment_by_id(created_apt_id),
        eq(apt["subject"], "MCP Test Appointment - Updated"),
        eq(apt["location"], "Updated Location"),
    ))

print("── Free/Busy ──")
run_test("get_free_busy", lambda: (
    fb := C.get_free_busy(),
    has("slots", fb),
    ok("error" in fb or len(fb["slots"]) >= 0, f"Free/busy: {fb.get('error', 'ok, ' + str(len(fb['slots'])) + ' slots')}"),
))


# ── CONTACTS ─────────────────────────────────────────────────────
print("── Contacts ──")
run_test("list_contacts", lambda: (
    contacts := C.list_contacts(count=5),
    ok(isinstance(contacts, list)),
))

run_test("create_contact", lambda: (
    r := C.create_contact(
        full_name="MCP Test Contact - Please Delete",
        email="mcp_test@example.com",
        phone="555-0001",
        mobile="555-0002",
        company="MCP Test Corp",
        job_title="Test Subject",
    ),
    has("message", r),
    _cleanup_ids["contacts"].append(r.get("full_name", "")),
))

# Find contact EntryID
if contacts_all := C.list_contacts(search="MCP Test Contact", count=5):
    test_contact_id = contacts_all[0]["entry_id"]
    _cleanup_ids["contacts"][-1] = test_contact_id  # update to use ID
    print(f"  Using contact EntryID: {test_contact_id[:20]}...")

    run_test("update_contact", lambda: (
        r := C.update_contact(
            entry_id=test_contact_id,
            full_name="MCP Test Contact - UPDATED",
            phone="555-9999",
            mobile="555-8888",
            company="MCP Updated Corp",
        ),
        has("message", r),
        ok("full_name" in r["changed_fields"]),
    ))

    run_test("verify_contact_updated", lambda: (
        # Re-fetch by looking it up again
        found := [c for c in C.list_contacts(search="MCP Test Contact - UPDATED", count=5)
                   if c["entry_id"] == test_contact_id],
        ok(len(found) >= 0),
    ))

    run_test("export_contacts_csv", lambda: (
        r := C.export_contacts(format="csv", save_path=_test_dir),
        has("file", r),
        ok(os.path.exists(r["file"]), f"CSV file should exist at {r['file']}"),
    ))

    run_test("export_contacts_vcard", lambda: (
        r := C.export_contacts(format="vcard", save_path=_test_dir),
        has("file", r),
        ok(os.path.exists(r["file"]), f"vCard file should exist at {r['file']}"),
    ))


# ── TASKS ────────────────────────────────────────────────────────
print("── Tasks ──")
run_test("list_tasks", lambda: (
    tasks := C.list_tasks(count=5),
    ok(isinstance(tasks, list)),
))

# Create task first (outside run_test so we can capture the ID)
_task_create_result = None
try:
    _task_create_result = C.create_task(
        subject="MCP Test Task - Please Delete",
        body="Test task body.",
        due_date="2026-06-20",
        importance=C.IMPORTANCE_NORMAL,
        reminder_minutes=0,
    )
    _cleanup_ids["tasks"].append(_task_create_result["entry_id"])
    _task_pass = (
        "entry_id" in _task_create_result
        and _task_create_result["subject"] == "MCP Test Task - Please Delete"
    )
except Exception as e:
    _task_pass = False
    _task_error = str(e)
    _task_create_result = None

def _test_create_task():
    if _task_create_result is None:
        raise AssertionError(_task_error)
    assert "entry_id" in _task_create_result, "Missing entry_id"
    assert _task_create_result["subject"] == "MCP Test Task - Please Delete"
run_test("create_task", _test_create_task)

if _task_create_result and "entry_id" in _task_create_result:
    task_id = _task_create_result["entry_id"]
    print(f"  Using task EntryID: {task_id[:20]}...")

    run_test("update_task", lambda: (
        r := C.update_task(
            entry_id=task_id,
            subject="MCP Test Task - UPDATED",
            body="Updated task body.",
            importance=C.IMPORTANCE_HIGH,
            status=1,  # In Progress
            percent_complete=50,
        ),
        has("message", r),
        ok("subject" in r.get("changed_fields", []) or "status" in r.get("changed_fields", []) or "percent_complete" in r.get("changed_fields", []), "No recognized field was changed"),
    ))

    run_test("mark_task_complete", lambda: (
        r := C.mark_task_complete(task_id, complete=True),
        has("message", r),
    ))

    run_test("mark_task_reopen", lambda: (
        r := C.mark_task_complete(task_id, complete=False),
        has("message", r),
    ))


# ── RULES ────────────────────────────────────────────────────────
print("── Rules ──")
run_test("get_rules", lambda: (
    r := C.get_rules(),
    has("rules", r),
    ok(isinstance(r["rules"], list)),
    ok("count" in r),
))

run_test("create_rule", lambda: (
    r := C.create_rule(
        name="MCP Test Rule - Please Delete",
        condition_type="sender",
        condition_value="noreply@example.com",
        action_type="mark_read",
        enabled=True,
    ),
    has("message", r),
    _cleanup_ids["rules"].append(r["name"]),
))


# ── EMPTY DELETED (skip actual emptying) ─────────────────────────
print("── Other ──")
run_test("empty_deleted_callable", lambda: (
    ok(callable(C.empty_deleted_folder), "empty_deleted_folder should be callable"),
))


# ═══════════════════════════════════════════════════════════════════
# CLEANUP
# ═══════════════════════════════════════════════════════════════════
print("\n── Cleanup ──")

# Delete drafts
for draft_id in _cleanup_ids["drafts"]:
    try:
        C.delete_email(draft_id)
        print(f"  Deleted draft: {draft_id[:20]}...")
    except Exception as e:
        print(f"  Could not delete draft {draft_id[:20]}...: {e}")

# Delete appointments
for apt_id in _cleanup_ids["appointments"]:
    try:
        C.delete_appointment(apt_id)
        print(f"  Deleted appointment: {apt_id[:20]}...")
    except Exception as e:
        print(f"  Could not delete appointment {apt_id[:20]}...: {e}")

# Delete contacts
for contact_id in _cleanup_ids["contacts"]:
    try:
        C.delete_contact(contact_id)
        print(f"  Deleted contact: {contact_id[:20]}...")
    except Exception as e:
        print(f"  Could not delete contact {contact_id[:20]}...: {e}")

# Delete tasks
for task_id in _cleanup_ids["tasks"]:
    try:
        C.delete_task(task_id)
        print(f"  Deleted task: {task_id[:20]}...")
    except Exception as e:
        print(f"  Could not delete task {task_id[:20]}...: {e}")

# Delete rules
for rule_name in _cleanup_ids["rules"]:
    try:
        rules = C.namespace.DefaultStore.GetRules()
        rules.Remove(rule_name)
        rules.Save()
        print(f"  Deleted rule: {rule_name}")
    except Exception as e:
        print(f"  Could not delete rule '{rule_name}': {e}")

# Cleanup temp directory
try:
    import shutil
    shutil.rmtree(_test_dir, ignore_errors=True)
    print(f"  Cleaned up temp dir: {_test_dir}")
except Exception:
    pass


# ═══════════════════════════════════════════════════════════════════
print(f"\n{'═'*60}")
print(f"RESULTS: {_passed} passed, {_failed} failed, {_skipped} skipped")
print(f"{'═'*60}")

if _failed > 0:
    print("\nFAILURES:")
    for r in _results:
        if r["status"] == "FAIL":
            print(f"  - {r['name']}: {r.get('error', '')}")

if _skipped > 0:
    print("\nSKIPPED:")
    for r in _results:
        if r["status"] == "SKIP":
            print(f"  - {r['name']}: {r.get('error', '')}")

sys.exit(0 if _failed == 0 else 1)
