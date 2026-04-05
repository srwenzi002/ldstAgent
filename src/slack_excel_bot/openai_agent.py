from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable
from zoneinfo import ZoneInfo

from openai import OpenAI

from slack_excel_bot.config import Settings
from slack_excel_bot.debug_trace import DebugTrace
from slack_excel_bot.excel_tools import ExcelToolService
from slack_excel_bot.tool_schemas import (
    AttendanceSheetInput,
    CalendarContextInput,
    ExpenseEvidenceAnalysisInput,
    PersonalExpenseSheetInput,
    TransportRouteBatchLookupInput,
    TransportRouteLookupInput,
    TransportSheetInput,
    openai_function_tool,
)


@dataclass
class AgentResult:
    text: str
    generated_files: list[dict[str, Any]] = field(default_factory=list)


class OpenAIExcelAgent:
    def __init__(self, settings: Settings, tool_service: ExcelToolService):
        self.settings = settings
        self.tool_service = tool_service
        self.tools = [
            openai_function_tool(
                "get_month_calendar_context",
                (
                    "日本の月次カレンダー情報を返します。"
                    "指定した year と month に対して、各日の曜日、土日かどうか、日本の祝日かどうか、祝日名を返します。"
                    "勤務表を作るときに、平日・土日・祝日の区別が必要なら、generate_attendance_sheet の前に必ずこのツールを使ってください。"
                    "月全体を『平日通常・土日休み』のように展開する場合、曜日や祝日の判断を自分で決めつけず、このツール結果を根拠にしてください。"
                ),
                CalendarContextInput,
            ),
            openai_function_tool(
                "generate_attendance_sheet",
                (
                    "日本向けの月次勤務表を作成します。"
                    "このツールを呼ぶときは、function call arguments に最終形に近い JSON を入れてください。"
                    "JSON をそのままユーザーへ見せないでください。"
                    "employee.department_code は 10/20/30/50/51/52/60/70 のいずれかにしてください。"
                    "days[].work_grade は 1/2/3/4 のみです。1=09:30-18:00, 2=09:00-17:30, 3=10:00-18:30, 4=10:30-19:00。"
                    "days[].leave_item_no は 1..15 のみです。"
                    "days[].work_grade の扱いには次のルールを守ってください。半休（days[].leave_item_no が 2 または 3）の場合は、days[].work_grade を必ず入力してください。"
                    "休日出勤の場合は、days[].work_grade を入力しないでください。"
                    "半休ではない休暇（全休・代休・振休など）の場合も、days[].work_grade を入力しないでください。"
                    "つまり work_grade を入れるのは通常勤務日と半休日のみです。"
                    "基本勤務時間が指定されている場合、通常勤務日には work_grade だけでなく clock_in と clock_out も必ず入力してください。"
                    "work_grade を入力した日には、必ずその日に対応する clock_in と clock_out も入力してください。"
                    "半休（days[].leave_item_no が 2 または 3）の場合は、work_grade に対応する定時を基準に、その半休区分に合わせた実勤務時間を clock_in と clock_out に入れてください。"
                    "つまり午前休なら午後勤務の開始・終了時刻、午後休なら午前勤務の開始・終了時刻を入れてください。"
                    "半休の時間境界はテンプレートの半休区切を使ってください。1 は 12:30、2 は 12:00、3 は 13:00、4 は 13:30 です。"
                    "ユーザーが例外として明示していない日本の祝日は、既定で休みとして扱ってください。"
                    "つまり『平日通常・土日休み』のような指定でも、祝日に特別な勤務指定がなければ work_grade と clock_in と clock_out を入れないでください。"
                    "schema にない項目は出力しないでください。"
                ),
                AttendanceSheetInput,
            ),
            openai_function_tool(
                "analyze_expense_evidence",
                (
                    "ユーザーの文章と画像から、証憑やスクリーンショットの種類を判定し、精算に必要な項目を抽出します。"
                    "このツールは解析専用で、Excel 生成や外部経路検索は行いません。"
                    "画像は交通系の履歴、領収書、請求書、レシートのいずれもありえます。"
                    "まず expense_type を transport / personal_expense / unknown のいずれかで判断し、根拠のある項目だけを埋めてください。"
                    "自信のない金額・日付・路線・店名は必ず null にしてください。推測は禁止です。"
                    "Suica/PASMO などの履歴画面なら、入・出・窓出・物販・定 などの原始イベントを transport_events にできるだけ残してください。"
                    "原始イベントから確実に精算対象の移動と判断できるものだけを transport_items にまとめてください。"
                    "物販を通常の乗車記録として扱わないでください。"
                    "定 も普通の入出場線索として扱って構いません。必要以上に特殊扱いせず、前後関係から移動を組み立ててください。"
                    "記録が複数画像にまたがる場合は、画像をまたいで判断して構いません。"
                    "transport_items が十分にそろっている場合は、top-level の単一交通項目だけで済ませないでください。"
                    "文章と画像の両方を使った場合、evidence_sources には text と image の両方を入れてください。"
                    "missing_fields には、次工程を止める不足項目を入れてください。"
                ),
                ExpenseEvidenceAnalysisInput,
            ),
            openai_function_tool(
                "lookup_transport_route_batch",
                (
                    "複数の交通明細について、経路候補と金額候補をまとめて照合します。"
                    "analyze_expense_evidence が複数の transport_items を抽出した場合は、まずこのツールを優先してください。"
                    "画像から読めた出発地・到着地・日付を Ekispert で照合し、各明細の候補を返します。"
                    "画像の金額と近い候補がある場合は matched_option を優先して使えます。"
                    "候補が曖昧なときは、いきなり Excel を作らずユーザー確認を優先してください。"
                    "round_trip_suggestions が返った場合は、同日・同額・往復関係の明細が見つかっています。resolved_items には既に往復統合済みの結果が入っているので、それを使って Excel を作成し、返信では『往復としてまとめました。片道2件に分けたい場合は教えてください』と案内してください。"
                ),
                TransportRouteBatchLookupInput,
            ),
            openai_function_tool(
                "lookup_transport_route_options",
                (
                    "交通経路と運賃の候補を検索します。"
                    "交通費精算をしたいが、正確な経路・路線名・金額が不足している場合は、このツールを優先してください。"
                    "日付は必ず絶対日付 YYYY-MM-DD にしてください。"
                    "route_from と route_to には、できるだけ駅名だけを入れてください。"
                    "候補が返ったら、自然文で 2〜3 件に絞って示し、すぐ Excel を作らずにユーザー確認を取ってください。"
                ),
                TransportRouteLookupInput,
            ),
            openai_function_tool(
                "generate_transport_sheet",
                (
                    "交通費精算表を作成します。"
                    "このツールを呼ぶときは、function call arguments に最終形に近い JSON を入れてください。"
                    "JSON をそのままユーザーへ見せないでください。"
                    "items は最大 18 件です。"
                    "items[].purpose は 営業活動, 客先作業, 研修・セミナー参加, 深夜帰宅, 接待関連, その他会社業務 のいずれかにしてください。"
                    "items[].transport_mode は 電車・バス または タクシー のみです。"
                    "items[].one_way_amount は片道金額です。is_round_trip=true の場合は帳票側で往復表示になります。"
                    "purpose が不明な場合は空欄でも構いません。プログラム側で 営業活動 を補います。"
                    "往復指定がない場合は空欄でも構いません。プログラム側で false を補います。"
                    "visit_place、route_line、receipt_no は不明なら空欄で構いません。"
                    "route_line を入れる場合は、換乗駅名は入れず、路線名だけを短く並べてください。"
                    "たとえば『青物横丁 -> 京急本線 -> 泉岳寺 -> 都営地下鉄浅草線 -> 押上 -> 京成押上線 -> 青砥』ではなく、『京急本線 -> 都営地下鉄浅草線 -> 京成押上線』のようにしてください。"
                    "schema にない項目は出力しないでください。"
                ),
                TransportSheetInput,
            ),
            openai_function_tool(
                "generate_personal_expense_sheet",
                (
                    "個人立替経費精算表を作成します。"
                    "このツールを呼ぶときは、function call arguments に最終形に近い JSON を入れてください。"
                    "JSON をそのままユーザーへ見せないでください。"
                    "items は最大 3 件です。"
                    "items[].purpose はテンプレートの許容値だけを使ってください。"
                    "items[].burden_department は正式な部署名のみです。"
                    "items[].project_code_name は許容される完全な案件コード名である必要があります。不明な場合は推測せず確認してください。"
                    "schema にない項目は出力しないでください。"
                ),
                PersonalExpenseSheetInput,
            ),
        ]
        self.handlers: dict[str, Callable[[dict[str, Any]], dict[str, Any]]] = {
            "get_month_calendar_context": self.tool_service.get_month_calendar_context,
            "generate_attendance_sheet": self.tool_service.generate_attendance_sheet,
            "analyze_expense_evidence": self.tool_service.analyze_expense_evidence,
            "lookup_transport_route_batch": self.tool_service.lookup_transport_route_batch,
            "lookup_transport_route_options": self.tool_service.lookup_transport_route_options,
            "generate_transport_sheet": self.tool_service.generate_transport_sheet,
            "generate_personal_expense_sheet": self.tool_service.generate_personal_expense_sheet,
        }

    def run(
        self,
        conversation_input: list[dict[str, Any]],
        trace: DebugTrace | None = None,
        status_callback: Callable[[str, list[str] | None], None] | None = None,
    ) -> AgentResult:
        client = OpenAI(api_key=self.settings.openai_api_key)
        today_jst = datetime.now(ZoneInfo("Asia/Tokyo")).date()
        instructions = (
            "あなたは Slack 上で動く、かわいくて頼れる事務アシスタントです。"
            "ユーザーとの会話は自然文で行い、ツールとのやり取りでは構造化 JSON を使ってください。"
            "基本の返信言語は日本語です。"
            "ただし、ユーザーの最新メッセージが中国語中心なら、中国語で返信してください。"
            "ユーザー向けの文体はやわらかく親しみやすく、読みやすい箇条書きを積極的に使ってください。"
            "各返信では、かわいい絵文字を 1〜3 個まで自然に使ってかまいません。"
            "Slack では GitHub Markdown が完全には使えないため、返信は Slack の mrkdwn に合わせてください。"
            "### 見出し、--- 区切り線、番号付き見出し、言語指定付きコードブロックは使わないでください。"
            "入力内容・反映内容・候補一覧・未反映項目などの一覧は、通常の bullet ではなく triple backticks のみを使ったコードブロックで包んで見やすく整形してください。"
            "コードブロックの開始と終了は ``` のみを使い、```text のような言語タグは付けないでください。"
            "コードブロックの外側では自然文で案内し、コードブロックの中では 1 行 1 項目で簡潔にまとめてください。"
            "強調が必要な場合は **太字** ではなく Slack 互換の *太字* を使ってください。"
            "内部推論や JSON、関数引数、ツールの生レスポンス、ファイルパスはユーザーへ見せないでください。"
            "画像の内容も理解できます。"
            "通常の雑談や簡単な質問なら、そのまま自然に答えてください。"
            "勤務表、交通費精算表、個人立替経費精算表の依頼なら、必要に応じてツールを使ってください。"
            "ツールを使う場合、JSON は function call arguments の中だけに置き、会話本文には絶対に出さないでください。"
            f"現在の日本時間の日付は {today_jst.isoformat()} です。今日・昨日・一昨日などの相対表現は、必ず絶対日付に直してからツールに渡してください。"
            "勤務表の依頼で、月全体の通常勤務を平日・土日・祝日ベースで展開する必要がある場合は、まず get_month_calendar_context を呼び出してください。"
            "平日・土日・祝日の判断を会話文だけで決めつけず、calendar tool の結果を見てから days を組み立ててください。"
            "交通費精算では purpose が不明なら 営業活動 を既定値として扱います。"
            "移動が1回だけ書かれていて往復か片道か不明な場合は、まず片道として扱います。"
            "電車・地下鉄・バスなどの公共交通はすべて 電車・バス に統一します。"
            "ユーザーが画像・スクリーンショット・領収書・レシートを送ったら、まず analyze_expense_evidence を使って交通費か個人経費か、あるいは判断不能かを見分けてください。"
            "交通カード履歴の画像では、最初に transport_events を整理し、その後で精算対象にできるものだけを transport_items にまとめてください。"
            "定 も入出場の一種として扱って構いません。不要に特殊扱いせず、前後関係から通常の移動として解釈してください。"
            "analyze_expense_evidence が transport_items を返したら、いきなり Excel を作らず、まず lookup_transport_route_batch で Ekispert 照合を行ってください。"
            "batch 結果に resolved_items がある場合は、それを優先して generate_transport_sheet に渡してください。"
            "batch 結果に round_trip_suggestions がある場合は、同日往復を既定で往復扱いに統合済みです。返信では『往復としてまとめました。片道2件に分けたい場合は教えてください』と案内してください。"
            "batch 結果に needs_confirmation が 1 件でもある場合は、表を先に作成したとしても、そのまま黙って終わらせないでください。"
            "その場合は返信内に必ず『今回まだ表に入れていない項目』または『確認したい項目』のコードブロックを作り、未反映の明細・理由・必要な確認内容を明記してください。"
            "lookup_transport_route_batch または lookup_transport_route_options の結果に prompt_reason='query_error' が含まれる、または error に『駅名が見つかりません』が含まれる場合、その明細をすぐ未反映にして最終回答を作ってはいけません。"
            "その場合は、error に含まれる駅名、元の画像テキスト、transport_events、transport_items、前後の記録、一般的な駅名表記を手がかりに、どの駅名が省略・略記・誤読・表記ゆれなのかを考えてください。"
            "高い確信で補完または正規化できる場合は、route_from または route_to の駅名だけを修正して、同じ lookup_transport_route_batch または lookup_transport_route_options をもう一度呼び出してください。"
            "query_error を受け取った直後に最終回答を作らず、まず再照会のための function call を優先してください。"
            "1回目の修正後も同じ種類の駅名エラーになった場合は、駅名だけを見直してもう1回だけ再試行して構いません。再試行は最大2回までです。"
            "修正してよいのは駅名だけです。日付・金額・方向・往復判定は推定で変更しないでください。高い確信がない場合は無理に修正しないでください。"
            "2回再試行しても駅名エラーが解決しない場合のみ、その明細を確認項目または未反映項目として案内してください。その際は、どの駅名をどう補完しようとしたかを短く添えてください。"
            "matched_option がある明細はその経路を優先しつつ、金額は画像やユーザー入力の IC 金額を優先して残してください。"
            "交通費精算表へ入れる route_line は、Ekispert の完全な経由駅列をそのまま使わず、路線名だけを短く並べてください。換乗駅名は不要です。"
            "候補が複数あって曖昧な場合は、Slack 互換の番号付きリストまたは番号付きコードブロックで示してユーザー確認を取ってください。"
            "画像またはユーザーが金額を明示している場合、one_way_amount はその金額を優先し、Ekispert の金額で上書きしないでください。"
            "transport_events のうち一部しか transport_items にできない場合は、取り込めたものだけで表を作って構いませんが、返信では未反映イベントも必ず明記してください。"
            "特に query_error、駅名省略、候補未確定、画像の途中切れなどで確定できなかった明細は、生成完了メッセージでも省略せず知らせてください。"
            "analyze_expense_evidence が transport で、travel_date・route_from・route_to・one_way_amount がそろっていれば generate_transport_sheet を使えます。"
            "transport だが金額や経路が足りなければ lookup_transport_route_options を使ってください。"
            "personal_expense と判断できたら、必要項目を集めて generate_personal_expense_sheet を使ってください。"
            "unknown の場合は勝手に表を作らず、交通費か個人経費かを確認してください。"
            "交通候補を提示するときは、番号付きのコードブロックを優先して、経路・片道金額・所要時間・乗換回数を見やすく示してください。"
            "画像の情報と文章が食い違う場合は、必ず先に確認してください。"
            "月日だけの交通画像で年が書かれていない場合、反証がなければ現在の日本年を優先してください。"
            "ファイル生成が成功したら、やさしい日本語または中国語で『できました』と案内し、Slack にファイルが届くことを伝えてください。"
            "交通費精算表の生成後に『今回、表に入力した明細』を案内する場合は、各行に区間・金額だけでなく、採用した route_line もできるだけ一緒に書いてください。"
            "例: 『京成上野 → 青砥 / 272円 / 京成本線』のように、最終的に採用した路線が分かる形にしてください。"
            "往復統合した場合も、可能なら『京急本線 → 都営地下鉄浅草線 → 京成押上線』のように、採用した route_line を添えてください。"
            "交通費精算表の生成後、未反映・保留・確認待ちのデータがあるなら、『今回、表に入力した明細』だけで終わらせず、続けて『今回まだ表に入れていない項目』を必ず出してください。"
        )

        if trace is not None:
            trace.write_section(
                "openai_request_1",
                {
                    "model": self.settings.openai_model,
                    "instructions": instructions,
                    "input": conversation_input,
                    "tools": self.tools,
                },
            )

        response = client.responses.create(
            model=self.settings.openai_model,
            instructions=instructions,
            input=conversation_input,
            tools=self.tools,
        )
        if trace is not None:
            trace.write_section("openai_response_1", response)
        generated_files: list[dict[str, Any]] = []

        for round_index in range(5):
            function_calls = [item for item in response.output if item.type == "function_call"]
            if not function_calls:
                return AgentResult(text=response.output_text.strip() or "はいっ、進めますね🌷", generated_files=generated_files)

            tool_outputs = []
            for call in function_calls:
                if status_callback is not None:
                    status_text, loading_messages = self._status_for_tool_name(call.name)
                    status_callback(status_text, loading_messages)
                handler = self.handlers[call.name]
                arguments = json.loads(call.arguments)
                result = handler(arguments)
                if result.get("output_path"):
                    generated_files.append(result)
                if trace is not None:
                    trace.write_section(
                        f"tool_call_{round_index + 1}_{call.name}",
                        {
                            "call_id": call.call_id,
                            "arguments": arguments,
                            "result": result,
                        },
                    )
                tool_outputs.append(
                    {
                        "type": "function_call_output",
                        "call_id": call.call_id,
                        "output": json.dumps(self._tool_result_summary(result), ensure_ascii=False),
                    }
                )

            if trace is not None:
                trace.write_section(
                    f"openai_followup_request_{round_index + 2}",
                    {
                        "model": self.settings.openai_model,
                        "previous_response_id": response.id,
                        "input": tool_outputs,
                        "tools": self.tools,
                    },
                )

            response = client.responses.create(
                model=self.settings.openai_model,
                instructions=instructions,
                previous_response_id=response.id,
                input=tool_outputs,
                tools=self.tools,
            )
            if trace is not None:
                trace.write_section(f"openai_followup_response_{round_index + 2}", response)

        return AgentResult(
            text="処理は進めましたが、このリクエストではツール呼び出し回数の上限に達しました…！🙏 入力を少し整理して、もう一度試してください。",
            generated_files=generated_files,
        )

    @staticmethod
    def _status_for_tool_name(tool_name: str) -> tuple[str, list[str]]:
        status_map = {
            "get_month_calendar_context": (
                "is checking calendar...",
                ["📆 月のカレンダーを確認中です", "🎌 祝日と曜日を整理しています"],
            ),
            "analyze_expense_evidence": (
                "is analyzing evidence...",
                ["🧾 画像の内容を確認中です", "✨ 証憑の項目を整理しています"],
            ),
            "lookup_transport_route_batch": (
                "is checking routes...",
                ["🚃 経路と運賃を確認中です", "💴 金額を照合しています"],
            ),
            "lookup_transport_route_options": (
                "is checking routes...",
                ["🚃 経路と運賃を確認中です", "🗂️ 候補をまとめています"],
            ),
            "generate_transport_sheet": (
                "is generating Excel...",
                ["📝 交通費精算表を作成中です", "📎 Excel を仕上げています"],
            ),
            "generate_personal_expense_sheet": (
                "is generating Excel...",
                ["🧮 経費精算表を作成中です", "📎 Excel を仕上げています"],
            ),
            "generate_attendance_sheet": (
                "is generating Excel...",
                ["📅 勤務データを整理中です", "📎 Excel を仕上げています"],
            ),
        }
        return status_map.get(tool_name, ("is thinking...", ["🌷 ご依頼を処理しています"]))

    @staticmethod
    def _tool_result_summary(result: dict[str, Any]) -> dict[str, Any]:
        if result.get("output_path"):
            return {
                "ok": True,
                "title": result.get("title"),
                "message": "Excel ファイルができました。現在の Slack スレッドへ自動でアップロードします。",
            }
        return result
