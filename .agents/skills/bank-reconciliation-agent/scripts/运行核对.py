"""Skill 入口。只在本次进程补充核对步骤，复用原程序的读表、剪枝及报表输出。"""
from bisect import bisect_left, bisect_right
from collections import Counter, defaultdict
from pathlib import Path
import sys

ENGINE = Path(__file__).resolve().with_name("engine")
if not (ENGINE / "银行核对入口.py").is_file():
    raise SystemExit("Skill 文件不完整，缺少 scripts/engine 中的核对程序。")
sys.path.insert(0, str(ENGINE))

import pandas as pd
import 银行核对入口 as entry
from balance import _closed_day_balances
from data_structures import BusinessEvent
from 核对期间 import coverage_dates, uncovered_dates
from llm_assistant import ONLINE_ALLOWED_FIELDS
from matcher import Matcher, select_non_conflicting_candidates
from 业务分组 import candidate_sort_key, has_business_conflict


class SkillMatcher(Matcher):
    skill_automatic = True
    DIRECT_CONFLICTS = frozenset({"业务编号或批次不同", "交易流水号不同", "账号", "卡号", "对方户名"})

    def _combination_has_conflict(self, evidence):
        return not self.DIRECT_CONFLICTS.isdisjoint(evidence["business_conflicts"])

    def _add_candidate(self, *args, **kwargs):
        if getattr(self, "scope_conflicts", ()):
            return None
        candidate = super()._add_candidate(*args, **kwargs)
        if candidate is None:
            return candidate
        if candidate.date_span_days > self.config.tolerance_days:
            candidate.evidence["selection_ineligible"] = True
            return candidate
        cache = getattr(self, "_purpose_results", {})
        existing = cache.get((candidate.bank_idxs, candidate.journal_idxs))
        if existing:
            candidate.evidence["business_conflicts"] = (existing,)
            self._score_existing_candidate(candidate)
            return candidate
        if candidate.metrics.total_diff_li:
            return candidate
        conflicts = candidate.evidence.get("business_conflicts", ())
        inferred = tuple(choice for choice in conflicts if choice not in self.DIRECT_CONFLICTS)
        if inferred:
            candidate.evidence["purpose_alerts"] = inferred
            candidate.evidence["business_conflicts"] = tuple(choice for choice in conflicts if choice in self.DIRECT_CONFLICTS)
            self._score_existing_candidate(candidate)
        return candidate

    def _purpose_rows(self, candidate):
        rows = {}
        for side, indices in (("bank", candidate.bank_idxs), ("journal", candidate.journal_idxs)):
            frame = self._frames()[side]
            rows[side] = [
                {"row": int(frame.at[i, "original_file_row"]) if "original_file_row" in frame else int(i),
                 "fields": {key: str(value).strip() for key, value in self._row_text_fields(frame, i).items()
                            if key in ONLINE_ALLOWED_FIELDS and str(value).strip()}}
                for i in indices
            ]
        return rows

    def _review_purposes(self, choices):
        """文字差异由 Agent 解释，只有核过原文的明确矛盾才能改变结论。"""
        exchange = getattr(self.llm_assistant, "exchange", None)
        if exchange is None:
            return
        cache = getattr(self, "_purpose_results", None)
        if cache is None:
            self._purpose_results = cache = {}
        def reject(candidate, reason):
            candidate.evidence.pop("verified_reconciliation", None)
            candidate.evidence["business_conflicts"] = (reason,)
            self._score_existing_candidate(candidate)

        requested = {}
        for candidate in sorted(choices, key=candidate_sort_key):
            key = candidate.bank_idxs, candidate.journal_idxs
            if key in cache:
                reject(candidate, cache[key])
                continue
            if candidate.evidence.get("purpose_alerts") and not has_business_conflict(candidate):
                requested.setdefault(key, []).append(candidate)
        distinct = list(requested.values())
        for start in range(0, len(distinct), 20):
            batch = {}
            for copies in distinct[start:start + 20]:
                candidate = copies[0]
                batch[candidate.candidate_id] = (copies, self._purpose_rows(candidate))
            payload = {
                "kind": "purpose",
                "candidates": [{"candidate_id": identity, **rows} for identity, (_, rows) in batch.items()],
            }
            answer, _ = exchange(payload)
            if answer is None:
                answer, _ = exchange(payload)
            if answer is None:
                continue
            verdicts = answer.get("verdicts")
            if not isinstance(verdicts, dict):
                self.llm_assistant.review_counts["未复核"] += 1
                continue
            valid = 0
            for identity, (copies, rows) in batch.items():
                verdict = verdicts.get(identity)
                if not isinstance(verdict, dict):
                    continue
                if verdict.get("status") == "unclear":
                    valid += 1
                    continue
                if verdict.get("status") not in {"compatible", "conflict"}:
                    continue
                basis = []
                for side in ("bank", "journal"):
                    citation = verdict.get(side)
                    if not isinstance(citation, dict) or not isinstance(citation.get("quote"), str) or not isinstance(citation.get("field"), str):
                        break
                    quote = citation["quote"].strip()
                    if not quote or not any(row["fields"].get(citation["field"]) == quote for row in rows[side]):
                        break
                    basis.append(f'{citation["field"]}“{quote}”')
                if len(basis) != 2:
                    continue
                valid += 1
                if verdict["status"] == "compatible":
                    continue
                conflict = "银行" + basis[0] + "与账上" + basis[1] + "用途冲突"
                cache[(copies[0].bank_idxs, copies[0].journal_idxs)] = conflict
                for candidate in copies:
                    reject(candidate, conflict)
            self.llm_assistant.review_counts["已完成" if valid == len(batch) else "未复核"] += 1

    def _lock_direct_matches(self):
        # 原表交易后余额可区分同额重复项。其余按整个日期容差窗口判断唯一性。
        gaps = self._balance_gaps()
        pairs = [c for c in self.candidates if c.match_type == "exact_1to1"
                 and c.evidence.get("verified_balance_match") and self._can_cover(c)
                 and not any(start <= self.bank.at[c.bank_idxs[0], "date"] <= end for start, end in gaps)]
        for candidate in self.candidates:
            if not candidate.evidence.get("verified_balance_match"):
                continue
            if any(start <= self.bank.at[candidate.bank_idxs[0], "date"] <= end for start, end in gaps):
                candidate.evidence.pop("verified_balance_match", None)
                candidate.evidence.pop("verified_reconciliation", None)
                self._score_existing_candidate(candidate)
        bank_count = Counter(c.bank_idxs for c in pairs)
        journal_count = Counter(c.journal_idxs for c in pairs)
        verified = [c for c in pairs if bank_count[c.bank_idxs] == journal_count[c.journal_idxs] == 1]
        self._review_purposes(verified)
        for candidate in verified:
            if not has_business_conflict(candidate):
                self._lock(candidate)

        amounts = {}
        for side, frame in (("bank", self.bank), ("journal", self.journal)):
            groups = defaultdict(list)
            for row in frame[["date", "amount_decimal"]].itertuples():
                if row.amount_decimal and row.Index not in self._locked_rows[side]:
                    groups[row.amount_decimal].append((pd.Timestamp(row.date).normalize(), int(row.Index)))
            amounts[side] = {amount: sorted(rows) for amount, rows in groups.items()}

        exact = {(c.bank_idxs[0], c.journal_idxs[0]): c for c in self.candidates if c.match_type == "exact_1to1"}
        window = pd.Timedelta(days=self.config.tolerance_days)

        def sole_neighbor(rows, days, day):
            start = bisect_left(days, day - window)
            end = bisect_right(days, day + window)
            return rows[start] if end - start == 1 else None

        pending = []
        for amount, bank in amounts["bank"].items():
            journal = amounts["journal"].get(amount, [])
            bank_days = [day for day, _ in bank]
            journal_days = [day for day, _ in journal]
            for bank_day, b in bank:
                other = sole_neighbor(journal, journal_days, bank_day)
                if other is None:
                    continue
                journal_day, j = other
                if sole_neighbor(bank, bank_days, journal_day) != (bank_day, b):
                    continue
                candidate = exact.get((b, j)) or self._add_candidate(
                    [b], [j], "exact_1to1" if bank_day == journal_day else "tolerance_date", "逐笔核对",
                )
                if candidate is not None and self._can_cover(candidate):
                    pending.append(candidate)
        self._review_purposes(pending)
        for candidate in pending:
            if not has_business_conflict(candidate):
                candidate.evidence["verified_reconciliation"] = "容差内双方同方向同金额各有一笔"
                self._lock(candidate)

    def _frames(self):
        return {"bank": self.bank, "journal": self.journal}

    def _months(self):
        return {side: set(frame["date"].dt.to_period("M")) for side, frame in self._frames().items()}

    def _balance_gaps(self):
        evidence = getattr(self, "coverage_evidence", {})
        gaps = [(start, end) for facts in evidence.values() for start, end, kind, _ in facts.get("periods", ())
                if kind == "不完整"]
        gaps.extend((day, day) for facts in evidence.values() for day in facts.get("balance_conflicts", {}))
        return gaps

    def _lock_balanced_windows(self):
        # 按日验证原始余额链。断点之前和之后分别计算。
        postings = {}
        for side, frame in self._frames().items():
            previous = getattr(self.overall_control, f"{side}_initial_balance", None)
            days = {}
            postings[side] = days
            if "balance" not in frame:
                return
            for day, rows in frame.groupby(frame["date"].dt.normalize(), sort=True):
                chain = _closed_day_balances(zip(rows["amount"], rows["balance"]), previous)
                days[day] = (chain, list(rows["amount_decimal"].items()))
                previous = chain[1] if chain else None

        months = self._months()
        shared_months = months["bank"] & months["journal"]
        balances = dict.fromkeys(postings)
        pending = {side: [] for side in postings}
        anchored = False
        already_locked = {side: set(values) for side, values in self._locked_rows.items()}
        before = len(self._locked_candidates)
        gaps = self._balance_gaps()
        last_day = None
        for day in sorted(postings["bank"].keys() | postings["journal"].keys()):
            previous_day, last_day = last_day, day
            if any(start <= day and (previous_day is None or previous_day <= end) for start, end in gaps):
                balances = dict.fromkeys(postings)
                anchored = False
                for entries in pending.values():
                    entries.clear()
                if any(start <= day <= end for start, end in gaps):
                    continue
            if day.to_period("M") not in shared_months:
                balances = dict.fromkeys(postings)
                anchored = False
                for entries in pending.values():
                    entries.clear()
                continue
            openings = balances.copy()
            discontinuity = False
            for side, days in postings.items():
                if day not in days:
                    continue
                chain, _ = days[day]
                if chain is None or (balances[side] is not None and balances[side] != chain[0]):
                    discontinuity = True
                openings[side], balances[side] = chain or (None, None)
            if discontinuity:
                anchored = False
                for entries in pending.values():
                    entries.clear()
            if openings["bank"] is not None and openings["bank"] == openings["journal"]:
                anchored = True
            if anchored:
                for side, days in postings.items():
                    if day in days:
                        pending[side].extend((i, value) for i, value in days[day][1]
                                             if i not in self._locked_rows[side])
            if balances["bank"] is not None and balances["bank"] == balances["journal"]:
                if anchored:
                    self.balance_windows_checked = True
                    self._match_window_remainder(pending["bank"], pending["journal"])
                anchored = True
                for entries in pending.values():
                    entries.clear()

        proposed = self._locked_candidates[before:]
        self._review_purposes(proposed)
        self._locked_candidates[before:] = [c for c in proposed if not has_business_conflict(c)]
        self._locked_rows = already_locked
        for candidate in self._locked_candidates[before:]:
            self._locked_rows["bank"].update(candidate.bank_idxs)
            self._locked_rows["journal"].update(candidate.journal_idxs)

    def _match_window_remainder(self, bank, journal):
        super()._match_window_remainder(bank, journal)
        # 同一段可核实的起止余额圈住剩余交易。逐方向核对组金额及用途。
        for direction in (1, -1):
            bank_rows = [(i, amount) for i, amount in bank
                         if amount * direction > 0 and i not in self._locked_rows["bank"]]
            journal_rows = [(i, amount) for i, amount in journal
                            if amount * direction > 0 and i not in self._locked_rows["journal"]]
            if (min(len(bank_rows), len(journal_rows)) == 0
                    or sum(amount for _, amount in bank_rows) != sum(amount for _, amount in journal_rows)):
                continue
            self._lock_balance_relationship([i for i, _ in bank_rows], [i for i, _ in journal_rows])

    def _dfs_one_to_many(self, source_type, target_type):
        for side, rows in self._locked_rows.items():
            self._combination_covered[side].update(rows)
        super()._dfs_one_to_many(source_type, target_type)

    @staticmethod
    def _can_cover(candidate):
        return (candidate.metrics.total_diff_li == 0
                and not candidate.evidence.get("selection_ineligible")
                and not candidate.evidence.get("total_only_without_boundary")
                and not has_business_conflict(candidate))

    def _prune_locked_candidates(self):
        self.candidates = [c for c in self.candidates
                           if not self._locked_rows["bank"].intersection(c.bank_idxs)
                           and not self._locked_rows["journal"].intersection(c.journal_idxs)]

    def _add_closed_candidate_groups(self):
        super()._add_closed_candidate_groups()
        exact = {}
        for candidate in sorted(self.candidates, key=candidate_sort_key):
            if self._can_cover(candidate):
                exact.setdefault((candidate.bank_idxs, candidate.journal_idxs), candidate)
        for group in self._group_ambiguous_candidates(list(exact.values())):
            if len(group) < 2:
                continue
            bank = sorted({i for c in group for i in c.bank_idxs})
            journal = sorted({i for c in group for i in c.journal_idxs})
            if not self._total_structure_matches(self.bank.loc[bank, "amount_decimal"].tolist(),
                                                 self.journal.loc[journal, "amount_decimal"].tolist()):
                continue
            completion = select_non_conflicting_candidates(group, should_stop=lambda: self.stopping)
            if sum(len(c.bank_idxs) + len(c.journal_idxs) for c in completion) != len(bank) + len(journal):
                continue
            if len(completion) == 1:
                continue
            candidate = self._add_candidate(bank, journal, "skill_closed_cover", "整组勾稽", resolves_full_group=True)
            if candidate is not None and not candidate.evidence.get("selection_ineligible"):
                self._review_purposes([candidate])
                if not has_business_conflict(candidate):
                    candidate.evidence["verified_reconciliation"] = "同一批原始记录可完整覆盖，收入和支出分别对上"
                    self._lock(candidate)
        self._prune_locked_candidates()

    def _refresh_candidate_ambiguity(self):
        super()._refresh_candidate_ambiguity()
        by_id = {c.candidate_id: c for c in self.candidates}
        for c in self.candidates:
            if not self._can_cover(c):
                continue
            alternatives = [identity for identity in c.evidence.get("alternative_candidate_ids", [])
                            if self._can_cover(by_id[identity])]
            c.evidence["alternative_candidate_ids"] = alternatives
            c.is_ambiguous = bool(alternatives)
            if not c.is_ambiguous:
                c.evidence["verified_reconciliation"] = "按现有原记录的日期、用途及收支核对，双方金额一致"
            else:
                c.evidence.pop("verified_reconciliation", None)

    def _group_ambiguous_candidates(self, candidates=None):
        if candidates is not None:
            return super()._group_ambiguous_candidates(candidates)
        groups = super()._group_ambiguous_candidates([c for c in self.candidates if self._can_cover(c)])
        return [group for group in groups if any(c.is_ambiguous for c in group)
                and len({(c.bank_idxs, c.journal_idxs) for c in group}) > 1]

    def _commit_selected_candidates(self):
        self.candidates = [c for c in self.candidates if not c.evidence.get("total_only_without_boundary")
                           and c.date_span_days <= self.config.tolerance_days]
        self._prune_locked_candidates()
        self._review_purposes(c for c in self.candidates if self._can_cover(c))
        self.row_conflicts = defaultdict(set)
        for c in self.candidates:
            known = getattr(self, "_purpose_results", {}).get((c.bank_idxs, c.journal_idxs))
            if known:
                c.evidence.pop("verified_reconciliation", None)
                c.evidence["business_conflicts"] = (known,)
            if has_business_conflict(c):
                reason = list(c.evidence.get("business_conflicts", ()))
                if c.text_evidence:
                    reason.extend(c.text_evidence.conflicting_fields)
                for side, rows in (("bank", c.bank_idxs), ("journal", c.journal_idxs)):
                    for row in rows:
                        self.row_conflicts[side, row].update(reason)
                c.evidence["selection_ineligible"] = True
        super()._commit_selected_candidates()
        for c in self.selected_candidates:
            c.processing_reason = "；".join(part for part in c.processing_reason.split("；")
                                          if not part.startswith(("综合可信度", "金额一致，业务依据不足")))
            alternatives = c.evidence.get("alternative_candidate_ids", []) if c.is_ambiguous else []
            if alternatives and not c.evidence.get("agent_confirmed"):
                c.processing_reason = "程序按原文、收付和日期比较后选定：" + self._original_rows(c)
        months = self._months()
        evidence = getattr(self, "coverage_evidence", {})
        supplied_months = {day.to_period("M") for day in coverage_dates(evidence)}
        calendar = months["bank"] | months["journal"] | supplied_months
        available = set(pd.period_range(min(calendar), max(calendar), freq="M")) if calendar else set()
        self.uncovered_months = {side: available - values for side, values in months.items()}
        observed = []
        for side, frame in self._frames().items():
            other = "journal" if side == "bank" else "bank"
            facts = evidence.get(side, {})
            for start, end, kind, source in facts.get("periods", []):
                if kind == "无交易" and frame.loc[frame["date"].dt.normalize().between(start, end), "amount_decimal"].ne(0).any():
                    raise ValueError(f"逐笔记录与原件无交易范围冲突：{source}")
            opposite_missing = uncovered_dates(frame["date"].dt.normalize(), self.uncovered_months[other],
                                               evidence.get(other, {}))
            unverified = frame.loc[~frame["matched"] & opposite_missing]
            for month, rows in unverified.groupby(unverified["date"].dt.to_period("M")):
                missing = "账上" if side == "bank" else "银行"
                observed.append(BusinessEvent(
                    event_id=f"本月无对侧逐笔-{side}-{month}", source=side, event_type="月度明细未覆盖",
                    row_idxs=tuple(rows.index), net_amount_li=int(rows["amount_decimal"].sum()),
                    relationship_formula=f"{month} {missing}一侧逐笔明细未覆盖对应日期",
                    evidence_basis=f"另一侧有 {len(rows)} 笔原记录，实际收付差额无法计算",
                ))
        grouped_rows = {i for event in observed if event.source == "bank" for i in event.row_idxs}
        grouped_rows.update(i for event in self.business_events if event.source == "bank" for i in event.row_idxs)
        observed.extend(self._bank_equal_cash_movements(grouped_rows))
        self.business_events[:0] = observed

    def _bank_equal_cash_movements(self, excluded):
        """把同日同对象等额收付的两笔放在同一个待查事项中。"""
        groups = defaultdict(list)
        for index, profile in self._business_rows["bank"].items():
            parties = profile["parties"].get("对方户名", ())
            amount = int(self.bank.at[index, "amount_decimal"])
            if len(parties) == 1 and amount:
                groups[profile["date"].normalize(), next(iter(parties)), abs(amount)].append(index)
        for (day, _, _), indices in groups.items():
            if len(indices) != 2 or any(i in excluded or self.bank.at[i, "matched"] for i in indices):
                continue
            records = self.bank.loc[indices]
            if records["amount_decimal"].sum() or "balance" not in records:
                continue
            if records["balance"].isna().any():
                continue
            first, second = records["amount"].tolist()
            b1, b2 = records["balance"].tolist()
            if b1 - first != b2 and b2 - second != b1:
                continue
            yield BusinessEvent(
                event_id="银行等额收付-" + "-".join(map(str, indices)), source="bank", event_type="银行同日等额收付",
                row_idxs=tuple(indices), net_amount_li=0,
                relationship_formula=f"{day:%Y-%m-%d} 同一对方收、付各 {abs(first):,.2f} 元，余额回到原值",
                evidence_basis="账上两笔均未找到对应，收入和支出差额分别保留",
            )


    def _original_rows(self, candidate):
        parts = []
        for side, indices, label in (("bank", candidate.bank_idxs, "银行"),
                                     ("journal", candidate.journal_idxs, "账上")):
            frame = self._frames()[side]
            rows = frame.loc[list(indices), "original_file_row"].tolist()
            head = "、".join(map(str, rows[:4]))
            end = f"等 {len(rows)} 行" if len(rows) > 4 else "行"
            parts.append(f"{label}原表第{head}{end}")
        return " 对 ".join(parts)


if __name__ == "__main__":
    raise SystemExit(entry.main(matcher_factory=SkillMatcher))
