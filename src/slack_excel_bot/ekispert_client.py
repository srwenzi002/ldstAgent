from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any

import httpx


class EkispertError(RuntimeError):
    """Raised when the Ekispert MCP server returns an error or invalid payload."""


@dataclass(frozen=True)
class RouteOption:
    option_id: str
    route_summary: str
    route_line: str
    one_way_amount: float
    total_minutes: int | None
    transfer_count: int | None

    def as_dict(self) -> dict[str, Any]:
        return {
            "option_id": self.option_id,
            "route_summary": self.route_summary,
            "route_line": self.route_line,
            "one_way_amount": self.one_way_amount,
            "total_minutes": self.total_minutes,
            "transfer_count": self.transfer_count,
        }


@dataclass(frozen=True)
class StationCandidate:
    station_code: str
    station_name: str
    station_yomi: str | None
    station_type: str
    station_type_detail: str | None
    prefecture_code: str | None
    prefecture_name: str | None

    def as_dict(self) -> dict[str, Any]:
        return {
            "station_code": self.station_code,
            "station_name": self.station_name,
            "station_yomi": self.station_yomi,
            "station_type": self.station_type,
            "station_type_detail": self.station_type_detail,
            "prefecture_code": self.prefecture_code,
            "prefecture_name": self.prefecture_name,
        }


class EkispertMcpClient:
    def __init__(
        self,
        access_key: str,
        *,
        endpoint: str = "https://api-mcp.ekispert.jp/mcp",
        timeout: float = 30.0,
    ):
        self.access_key = access_key.strip()
        self.endpoint = endpoint
        self.timeout = timeout

    def search_route_options(
        self,
        *,
        route_from: str,
        route_to: str,
        top_k: int = 3,
        travel_date: str | None = None,
    ) -> list[RouteOption]:
        if not self.access_key:
            raise EkispertError("Ekispert access key is not configured.")

        arguments: dict[str, Any] = {
            "viaList": f"{route_from}:{route_to}",
            "searchType": "plain",
            "sort": "price",
            "ticketSystemType": "ic",
            "preferredTicketOrder": "cheap",
        }
        if travel_date:
            arguments["date"] = int(travel_date.replace("-", ""))

        payload = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": "ekispert_api_search_routes",
                "arguments": arguments,
            },
        }

        with httpx.Client(timeout=self.timeout) as client:
            response = client.post(
                self.endpoint,
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream",
                    "MCP-Protocol-Version": "2025-06-18",
                    "ekispert-api-access-key": self.access_key,
                    "ekispert-api-response-format": "json",
                },
                json=payload,
            )
            response.raise_for_status()

        message = self._parse_mcp_response(response.text)
        result = message.get("result")
        if not isinstance(result, dict):
            raise EkispertError("Ekispert MCP response did not include a result object.")
        if result.get("isError"):
            raise EkispertError(EkispertMcpClient._extract_error_message(result))

        content = result.get("content")
        if not isinstance(content, list) or not content:
            raise EkispertError("Ekispert MCP response did not include route content.")

        first_content = content[0]
        if not isinstance(first_content, dict) or not isinstance(first_content.get("text"), str):
            raise EkispertError("Ekispert MCP response content was not valid JSON text.")

        body = json.loads(first_content["text"])
        return self._parse_route_options(body, top_k=top_k)

    def search_station_candidates(
        self,
        *,
        station_name: str,
        top_k: int = 5,
        prefecture_code: str | None = "13",
        match_type: str = "partial",
        station_type: str = "train",
    ) -> list[StationCandidate]:
        if not self.access_key:
            raise EkispertError("Ekispert access key is not configured.")

        params: dict[str, Any] = {
            "key": self.access_key,
            "name": station_name,
            "nameMatchType": match_type,
        }
        if station_type:
            params["type"] = station_type
        if prefecture_code:
            params["prefectureCode"] = prefecture_code

        with httpx.Client(timeout=self.timeout) as client:
            response = client.get("https://api.ekispert.jp/v1/json/station/light", params=params)
            response.raise_for_status()

        body = response.json()
        return self._parse_station_candidates(body, top_k=top_k)

    @staticmethod
    def _extract_error_message(result: dict[str, Any]) -> str:
        content = result.get("content")
        if isinstance(content, list):
            for item in content:
                if isinstance(item, dict) and isinstance(item.get("text"), str):
                    return item["text"]
        return "Ekispert MCP returned an unknown error."

    @staticmethod
    def _parse_mcp_response(raw_text: str) -> dict[str, Any]:
        data_lines = [line.removeprefix("data: ").strip() for line in raw_text.splitlines() if line.startswith("data: ")]
        if not data_lines:
            raise EkispertError("Ekispert MCP response did not contain any data frames.")
        return json.loads(data_lines[-1])

    @staticmethod
    def _parse_route_options(body: dict[str, Any], *, top_k: int) -> list[RouteOption]:
        result_set = body.get("ResultSet")
        if not isinstance(result_set, dict):
            raise EkispertError("Ekispert route payload did not contain ResultSet.")

        raw_courses = result_set.get("Course")
        if isinstance(raw_courses, dict):
            raw_courses = [raw_courses]
        if not isinstance(raw_courses, list) or not raw_courses:
            raise EkispertError("No route candidates were returned by Ekispert.")

        options: list[RouteOption] = []
        for index, course in enumerate(raw_courses[:top_k], start=1):
            if not isinstance(course, dict):
                continue
            route_line = EkispertMcpClient._extract_route_line(course)
            route_summary = EkispertMcpClient._extract_route_summary(course, route_line)
            one_way_amount = EkispertMcpClient._extract_one_way_amount(course)
            total_minutes = EkispertMcpClient._extract_total_minutes(course)
            transfer_count = EkispertMcpClient._extract_transfer_count(course)
            options.append(
                RouteOption(
                    option_id=str(index),
                    route_summary=route_summary,
                    route_line=route_line,
                    one_way_amount=one_way_amount,
                    total_minutes=total_minutes,
                    transfer_count=transfer_count,
                )
            )

        if not options:
            raise EkispertError("No valid route candidates could be parsed from Ekispert response.")
        return options

    @staticmethod
    def _parse_station_candidates(body: dict[str, Any], *, top_k: int) -> list[StationCandidate]:
        result_set = body.get("ResultSet")
        if not isinstance(result_set, dict):
            raise EkispertError("Ekispert station payload did not contain ResultSet.")

        raw_points = result_set.get("Point")
        if isinstance(raw_points, dict):
            raw_points = [raw_points]
        if not isinstance(raw_points, list) or not raw_points:
            return []

        candidates: list[StationCandidate] = []
        for point in raw_points:
            if not isinstance(point, dict):
                continue
            station = point.get("Station")
            if not isinstance(station, dict):
                continue
            station_code = str(station.get("code") or "").strip()
            station_name = str(station.get("Name") or "").strip()
            if not station_code or not station_name:
                continue

            station_type_value = station.get("Type")
            if isinstance(station_type_value, dict):
                station_type = str(station_type_value.get("text") or "").strip()
                station_type_detail = str(station_type_value.get("detail") or "").strip() or None
            else:
                station_type = str(station_type_value or "").strip()
                station_type_detail = None

            prefecture = point.get("Prefecture")
            prefecture_code = None
            prefecture_name = None
            if isinstance(prefecture, dict):
                prefecture_code = str(prefecture.get("code") or "").strip() or None
                prefecture_name = str(prefecture.get("Name") or "").strip() or None

            candidates.append(
                StationCandidate(
                    station_code=station_code,
                    station_name=station_name,
                    station_yomi=str(station.get("Yomi") or "").strip() or None,
                    station_type=station_type,
                    station_type_detail=station_type_detail,
                    prefecture_code=prefecture_code,
                    prefecture_name=prefecture_name,
                )
            )
            if len(candidates) >= top_k:
                break

        return candidates

    @staticmethod
    def _extract_route_line(course: dict[str, Any]) -> str:
        teiki = course.get("Teiki")
        if isinstance(teiki, dict):
            display_route = teiki.get("DisplayRoute")
            if isinstance(display_route, str) and display_route.strip():
                return display_route.replace("--", " -> ")

        route = course.get("Route")
        if not isinstance(route, dict):
            return ""
        lines = route.get("Line")
        if isinstance(lines, dict):
            lines = [lines]
        if not isinstance(lines, list):
            return ""
        names = [line.get("Name") for line in lines if isinstance(line, dict) and isinstance(line.get("Name"), str)]
        return " / ".join(names)

    @staticmethod
    def _extract_route_summary(course: dict[str, Any], route_line: str) -> str:
        route = course.get("Route")
        if not isinstance(route, dict):
            return route_line
        points = route.get("Point")
        if isinstance(points, dict):
            points = [points]
        if not isinstance(points, list) or not points:
            return route_line
        station_names: list[str] = []
        for point in points:
            if not isinstance(point, dict):
                continue
            station = point.get("Station")
            if isinstance(station, dict) and isinstance(station.get("Name"), str):
                station_names.append(station["Name"])
        if station_names:
            return " -> ".join(station_names)
        return route_line

    @staticmethod
    def _extract_one_way_amount(course: dict[str, Any]) -> float:
        prices = course.get("Price")
        if isinstance(prices, dict):
            prices = [prices]
        if not isinstance(prices, list):
            raise EkispertError("Route candidate did not contain pricing data.")

        for price in prices:
            if not isinstance(price, dict):
                continue
            if price.get("kind") == "FareSummary":
                value = price.get("Oneway")
                if value is not None:
                    return float(value)

        raise EkispertError("Route candidate did not include a FareSummary one-way amount.")

    @staticmethod
    def _extract_total_minutes(course: dict[str, Any]) -> int | None:
        route = course.get("Route")
        if not isinstance(route, dict):
            return None

        total = 0
        found = False
        for key in ("timeOnBoard", "timeOther", "timeWalk"):
            value = route.get(key)
            if value is None:
                continue
            total += int(value)
            found = True
        return total if found else None

    @staticmethod
    def _extract_transfer_count(course: dict[str, Any]) -> int | None:
        route = course.get("Route")
        if not isinstance(route, dict):
            return None
        value = route.get("transferCount")
        return int(value) if value is not None else None
