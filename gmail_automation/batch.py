from __future__ import annotations

import uuid
from datetime import datetime

from .models import ConfirmationRow, SendBatch, SendConfig


class BatchPlanner:
    def plan(self, rows: list[ConfirmationRow], config: SendConfig) -> list[SendBatch]:
        eligible = [row for row in rows if self._eligible(row)]
        eligible = eligible[: config.daily_send_limit]
        batches: list[SendBatch] = []
        for index in range(0, len(eligible), config.batch_size):
            batch_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
            batches.append(SendBatch(batch_id=batch_id, rows=eligible[index : index + config.batch_size]))
        return batches

    def _eligible(self, row: ConfirmationRow) -> bool:
        state = row.state
        if state.ready_to_send != "Y" or state.verification_status != "Passed":
            return False
        if state.main_sent == "Y" or state.retry_locked == "Y":
            return False
        if state.next_retry_at:
            try:
                if datetime.strptime(state.next_retry_at, "%Y-%m-%d %H:%M:%S") > datetime.now():
                    return False
            except ValueError:
                return False
        return True
