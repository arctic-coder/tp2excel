"""TrainingPeaks API client."""

import requests
from datetime import date, timedelta

TP_API_URL = "https://tpapi.trainingpeaks.com"

WORKOUT_TYPE_MAP = {
    1: "Swim",
    2: "Bike",
    3: "Run",
    4: "Brick",
    5: "CrossTrain",
    7: "DayOff",
    8: "MTB",
    9: "Weight",
    11: "XC-Ski",
    12: "Rowing",
    13: "Walk",
    100: "Other",
}


class TrainingPeaksClient:
    def __init__(self, auth_cookie: str, plan_days_shift: int = 0):
        self._auth_cookie = auth_cookie
        self._plan_days_shift = plan_days_shift
        self._token: str | None = None
        self._user_id: str | None = None

    def _get_token(self) -> str:
        if self._token:
            return self._token
        resp = requests.get(
            f"{TP_API_URL}/users/v3/token",
            headers={"Cookie": self._auth_cookie},
        )
        resp.raise_for_status()
        self._token = resp.json()["token"]["access_token"]
        return self._token

    def _auth_headers(self) -> dict:
        return {"Authorization": f"Bearer {self._get_token()}"}

    def _get_user_id(self) -> str:
        if self._user_id:
            return self._user_id
        resp = requests.get(
            f"{TP_API_URL}/users/v3/user",
            headers=self._auth_headers(),
        )
        resp.raise_for_status()
        self._user_id = str(resp.json()["user"]["userId"])
        return self._user_id

    def get_plans(self) -> list[dict]:
        resp = requests.get(
            f"{TP_API_URL}/plans/v1/plans",
            headers=self._auth_headers(),
        )
        resp.raise_for_status()
        return resp.json()

    def get_workouts_from_plan(self, plan_id: str) -> list[dict]:
        user_id = self._get_user_id()
        plan_start = _this_monday() + timedelta(days=self._plan_days_shift)

        apply_resp = requests.post(
            f"{TP_API_URL}/plans/v1/commands/applyplan",
            json=[{"athleteId": user_id, "planId": plan_id,
                   "targetDate": plan_start.isoformat(), "startType": "1"}],
            headers=self._auth_headers(),
        )
        apply_resp.raise_for_status()
        applied = apply_resp.json()[0]
        applied_plan_id = applied["appliedPlanId"]
        end_date = date.fromisoformat(applied["endDate"][:10])

        try:
            workouts = self._get_calendar_workouts(user_id, plan_start, end_date)
            shift_back = timedelta(days=self._plan_days_shift)
            for w in workouts:
                d = date.fromisoformat(w["workoutDay"][:10])
                w["workoutDay"] = (d - shift_back).isoformat()
            return workouts
        finally:
            requests.post(
                f"{TP_API_URL}/plans/v1/commands/removeplan",
                json={"appliedPlanId": applied_plan_id},
                headers=self._auth_headers(),
            )

    def _get_calendar_workouts(self, user_id: str, start: date, end: date) -> list[dict]:
        resp = requests.get(
            f"{TP_API_URL}/fitness/v6/athletes/{user_id}/workouts"
            f"/{start.isoformat()}/{end.isoformat()}",
            headers=self._auth_headers(),
        )
        resp.raise_for_status()
        return resp.json()


def _this_monday() -> date:
    today = date.today()
    return today - timedelta(days=today.weekday())


def parse_workout_row(workout: dict, plan_start: date) -> dict:
    workout_date = date.fromisoformat(workout.get("workoutDay", "")[:10])
    delta = (workout_date - plan_start).days
    week = delta // 7 + 1
    day_name = workout_date.strftime("%A")

    duration_hours = workout.get("totalTimePlanned") or 0.0
    duration_min = round(duration_hours * 60)

    description = (workout.get("description") or "").strip()
    coach_comments = (workout.get("coachComments") or "").strip()
    if coach_comments:
        description = f"{description}\n---\n{coach_comments}".strip()

    type_id = workout.get("workoutTypeValueId")
    workout_type = WORKOUT_TYPE_MAP.get(type_id, "Unknown") if type_id else "Unknown"

    return {
        "week": week,
        "day": day_name,
        "type": workout_type,
        "name": (workout.get("title") or "Workout").strip(),
        "duration_min": duration_min if duration_min > 0 else "",
        "tss": workout.get("tssPlanned") or "",
        "description": description,
    }
