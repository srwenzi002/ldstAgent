from pathlib import Path

from slack_excel_bot.thread_draft_store import ThreadDraftStore


def test_thread_draft_store_replaces_same_template_and_keeps_history(tmp_path: Path) -> None:
    store = ThreadDraftStore(tmp_path)
    thread_ts = "123.456"

    first = store.upsert_draft(
        thread_ts=thread_ts,
        template_type="transport",
        mode="merge",
        draft_patch={
            "employee": {"employee_id": "0001", "name": "山田太郎", "department": "開発本部", "department_code": "50"},
            "items": [
                {
                    "travel_date": "2026-04-02",
                    "purpose": "営業活動",
                    "transport_mode": "電車・バス",
                    "route_from": "新橋",
                    "route_to": "新高円寺",
                    "one_way_amount": 252,
                    "is_round_trip": False,
                }
            ],
        },
        pending_questions=["経路を確認してください"],
    )
    assert first.replaced_previous is False
    assert first.snapshot["drafts_by_template"]["transport"]["status"] == "waiting_confirmation"

    store.record_file_generated(
        thread_ts=thread_ts,
        template_type="transport",
        generated_file={"template_id": "transport_jp_leadingsoft_v1", "title": "交通費", "output_path": "/tmp/a.xlsx"},
        canonical_state=first.snapshot["drafts_by_template"]["transport"]["canonical_state"],
    )

    second = store.upsert_draft(
        thread_ts=thread_ts,
        template_type="transport",
        mode="replace",
        draft_patch={
            "employee": {"employee_id": "0002", "name": "李四", "department": "営業部", "department_code": "20"},
            "items": [
                {
                    "travel_date": "2026-04-03",
                    "purpose": "営業活動",
                    "transport_mode": "電車・バス",
                    "route_from": "品川",
                    "route_to": "新宿",
                    "one_way_amount": 208,
                    "is_round_trip": False,
                }
            ],
        },
        pending_questions=[],
    )

    draft = second.snapshot["drafts_by_template"]["transport"]
    assert second.replaced_previous is True
    assert draft["canonical_state"]["employee"]["employee_id"] == "0002"
    assert draft["latest_generated_file"] is None
    assert len(draft["generation_history"]) == 1


def test_thread_draft_store_builds_context_summary(tmp_path: Path) -> None:
    store = ThreadDraftStore(tmp_path)
    thread_ts = "abc"

    store.upsert_draft(
        thread_ts=thread_ts,
        template_type="personal_expense",
        mode="merge",
        draft_patch={
            "employee": {"employee_id": "1011", "name": "シュウ", "department": "ソリューション開発部", "department_code": "51"},
            "items": [{"expense_date": "2026-02-16", "payee_name": "麻辣大学上野店"}],
        },
        pending_questions=["用途を確認してください"],
    )

    summary = store.build_context_summary(thread_ts)

    assert "personal_expense" in summary
    assert "用途を確認してください" in summary
    assert "麻辣大学上野店" in summary
