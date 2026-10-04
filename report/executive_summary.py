"""
Executive Summary sheet generator for the AWS Weekly BAU Report.

Produces auto-generated textual summaries, key highlights, and a
traffic-light infrastructure health overview derived from collected data.
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from report.styles import ReportStyles


class ExecutiveSummarySheet:
    """Generates the Executive Summary tab."""

    SHEET_NAME = "Executive Summary"
    TOTAL_COLS = 8

    def generate(self, wb, all_data):
        """
        Create the Executive Summary sheet.

        Parameters
        ----------
        wb : openpyxl.Workbook
        all_data : dict
            Complete collected_data dictionary from all collectors.

        Returns
        -------
        openpyxl.worksheet.worksheet.Worksheet
        """
        ws = wb.create_sheet(self.SHEET_NAME)
        ws.sheet_properties.tabColor = ReportStyles.DARK_BLUE
        styles = ReportStyles

        # Column widths
        for c in range(1, self.TOTAL_COLS + 1):
            ws.column_dimensions[get_column_letter(c)].width = 18

        row = 1

        # ── Title bar ────────────────────────────────────────────────
        styles.apply_title_bar(ws, row, "Executive Summary", self.TOTAL_COLS)
        row += 2

        # ── Cost Overview ─────────────────────────────────────────────
        cost = all_data.get("cost", {})
        row = self._write_cost_overview(ws, row, cost)
        row += 1

        # ── Infrastructure Overview ───────────────────────────────────
        ec2_list = all_data.get("ec2", [])
        row = self._write_infra_overview(ws, row, ec2_list)
        row += 1

        # ── Service Summary ───────────────────────────────────────────
        row = self._write_service_summary(ws, row, all_data)
        row += 1

        # ── Health Traffic Lights ─────────────────────────────────────
        row = self._write_health_lights(ws, row, all_data)

        ws.sheet_view.showGridLines = False
        return ws

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _section(self, ws, row, title):
        ReportStyles.apply_section_header(ws, row, title, self.TOTAL_COLS)
        return row + 1

    def _kv_row(self, ws, row, label, value, fmt=None):
        """Write a label: value pair across columns B–C."""
        lbl = ws.cell(row=row, column=2, value=label)
        lbl.font = ReportStyles.FONT_LABEL
        lbl.alignment = Alignment(horizontal="right", vertical="center")

        val = ws.cell(row=row, column=3, value=value)
        val.font = ReportStyles.FONT_VALUE
        val.alignment = Alignment(horizontal="left", vertical="center")
        if fmt:
            val.number_format = fmt
        ws.row_dimensions[row].height = 22
        return row + 1

    def _text_row(self, ws, row, text, cols=None):
        """Write a full-width narrative text row."""
        cols = cols or self.TOTAL_COLS
        ws.merge_cells(f"B{row}:{get_column_letter(cols)}{row}")
        cell = ws.cell(row=row, column=2, value=text)
        cell.font = ReportStyles.FONT_VALUE
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[row].height = 28
        return row + 1

    # ── Cost Overview ─────────────────────────────────────────────────

    def _write_cost_overview(self, ws, row, cost):
        row = self._section(ws, row, "💰 Cost Overview")
        row += 1

        current = cost.get("current_week_total", 0)
        previous = cost.get("previous_week_total", 0)
        diff = cost.get("difference", 0)
        pct = cost.get("pct_change", 0)
        avg_daily = cost.get("avg_daily_cost", 0)

        row = self._kv_row(ws, row, "Current Week Cost", f"${current:,.2f}")
        row = self._kv_row(ws, row, "Previous Week Cost", f"${previous:,.2f}")

        # Difference with color
        diff_str = f"${abs(diff):,.2f}"
        if diff > 0:
            diff_str = f"+{diff_str}  (↑ {abs(pct):.1f}%)"
        elif diff < 0:
            diff_str = f"-{diff_str}  (↓ {abs(pct):.1f}%)"
        else:
            diff_str = f"{diff_str}  (No change)"
        row = self._kv_row(ws, row, "Week-over-Week Change", diff_str)

        row = self._kv_row(ws, row, "Average Daily Cost", f"${avg_daily:,.2f}")

        # Narrative
        if pct > ReportStyles.COST_INCREASE_CRITICAL:
            narrative = (
                f"⚠️  Cost increased by {abs(pct):.1f}% week-over-week, which exceeds "
                f"the critical threshold of {ReportStyles.COST_INCREASE_CRITICAL}%. "
                "Immediate review of cost drivers is recommended."
            )
        elif pct > ReportStyles.COST_INCREASE_WARNING:
            narrative = (
                f"⚡  Cost increased by {abs(pct):.1f}% week-over-week. "
                "This is above the warning threshold — monitor closely."
            )
        elif diff < 0:
            narrative = (
                f"✅  Cost decreased by {abs(pct):.1f}% compared to the previous week. "
                "Optimisation efforts are yielding results."
            )
        else:
            narrative = "✅  Cost is stable with minimal week-over-week change."

        row += 1
        row = self._text_row(ws, row, narrative)
        return row

    # ── Infrastructure Overview ───────────────────────────────────────

    def _write_infra_overview(self, ws, row, ec2_list):
        row = self._section(ws, row, "🖥️ Infrastructure Overview")
        row += 1

        total = len(ec2_list)
        running = sum(1 for i in ec2_list if str(i.get("state", "")).lower() == "running")
        stopped = sum(1 for i in ec2_list if str(i.get("state", "")).lower() == "stopped")
        other = total - running - stopped

        row = self._kv_row(ws, row, "Total EC2 Instances", str(total))
        row = self._kv_row(ws, row, "Running", str(running))
        row = self._kv_row(ws, row, "Stopped", str(stopped))
        if other:
            row = self._kv_row(ws, row, "Other States", str(other))

        # CPU / Memory highlights
        cpus = [i.get("cpu_avg") for i in ec2_list if i.get("cpu_avg") is not None]
        cpu_maxs = [i.get("cpu_max") for i in ec2_list if i.get("cpu_max") is not None]
        mems = [i.get("memory_avg") for i in ec2_list if i.get("memory_avg") is not None]
        mem_maxs = [i.get("memory_max") for i in ec2_list if i.get("memory_max") is not None]

        if cpus:
            avg_cpu = sum(cpus) / len(cpus)
            max_cpu = max(cpu_maxs) if cpu_maxs else max(cpus)
            row = self._kv_row(ws, row, "Avg CPU Utilisation", f"{avg_cpu:.1f}%")
            row = self._kv_row(ws, row, "Peak CPU Utilisation", f"{max_cpu:.1f}%")
        if mems:
            avg_mem = sum(mems) / len(mems)
            max_mem = max(mem_maxs) if mem_maxs else max(mems)
            row = self._kv_row(ws, row, "Avg Memory Utilisation", f"{avg_mem:.1f}%")
            row = self._kv_row(ws, row, "Peak Memory Utilisation", f"{max_mem:.1f}%")

        return row

    # ── Service Summary ───────────────────────────────────────────────

    def _write_service_summary(self, ws, row, all_data):
        row = self._section(ws, row, "📋 Service Summary")
        row += 1

        items = [
            ("EC2 Instances", len(all_data.get("ec2", []))),
            ("Load Balancers (ALB)", len(all_data.get("elb", []))),
            ("WAF Web ACLs", len(all_data.get("waf", []))),
            ("S3 Buckets", len(all_data.get("s3", []))),
            ("RDS Instances", len(all_data.get("rds", []))),
        ]
        for label, count in items:
            row = self._kv_row(ws, row, label, str(count))

        # Security findings (Inspector + GuardDuty) — headline security data
        inspector = all_data.get("inspector") or {}
        guardduty = all_data.get("guardduty") or {}
        insp_total = int(inspector.get("total_findings", 0) or 0)
        gd_total = int(guardduty.get("total_findings", 0) or 0)
        insp_counts = inspector.get("severity_counts", {}) or {}
        insp_crit_high = int(insp_counts.get("CRITICAL", 0) or 0) + int(
            insp_counts.get("HIGH", 0) or 0
        )
        gd_high = int((guardduty.get("severity_counts", {}) or {}).get("HIGH", 0) or 0)

        row = self._kv_row(
            ws, row,
            "Security Findings",
            f"Inspector {insp_total} (Crit/High {insp_crit_high})  |  "
            f"GuardDuty {gd_total} (High {gd_high})",
        )
        return row

    # ── Health Traffic Lights ─────────────────────────────────────────

    def _write_health_lights(self, ws, row, all_data):
        row = self._section(ws, row, "🚦 Health Status")
        row += 1

        headers = ["Category", "Status", "Details"]
        for idx, h in enumerate(headers, start=2):
            cell = ws.cell(row=row, column=idx, value=h)
            cell.font = ReportStyles.FONT_HEADER
            cell.fill = ReportStyles.FILL_DARK_BLUE
            cell.alignment = ReportStyles.ALIGN_CENTER
            cell.border = ReportStyles.THIN_BORDER
        row += 1

        checks = self._compute_health(all_data)
        for i, (category, status, detail) in enumerate(checks):
            even = i % 2 == 0
            fill = ReportStyles.FILL_LIGHT_GRAY if even else ReportStyles.FILL_WHITE

            cat_cell = ws.cell(row=row, column=2, value=category)
            cat_cell.font = ReportStyles.FONT_DATA_BOLD
            cat_cell.fill = fill
            cat_cell.border = ReportStyles.THIN_BORDER
            cat_cell.alignment = ReportStyles.ALIGN_LEFT

            status_cell = ws.cell(row=row, column=3, value=status)
            status_cell.font = ReportStyles.FONT_DATA_BOLD
            status_cell.alignment = ReportStyles.ALIGN_CENTER
            status_cell.border = ReportStyles.THIN_BORDER
            if "GREEN" in status.upper() or "HEALTHY" in status.upper():
                status_cell.fill = ReportStyles.FILL_GREEN
                status_cell.font = Font(name="Calibri", size=10, bold=True, color=ReportStyles.GREEN_FG)
            elif "YELLOW" in status.upper() or "WARNING" in status.upper():
                status_cell.fill = ReportStyles.FILL_YELLOW
                status_cell.font = Font(name="Calibri", size=10, bold=True, color=ReportStyles.YELLOW_FG)
            elif "RED" in status.upper() or "CRITICAL" in status.upper():
                status_cell.fill = ReportStyles.FILL_RED
                status_cell.font = Font(name="Calibri", size=10, bold=True, color=ReportStyles.RED_FG)
            else:
                status_cell.fill = fill

            det_cell = ws.cell(row=row, column=4, value=detail)
            det_cell.font = ReportStyles.FONT_DATA
            det_cell.fill = fill
            det_cell.border = ReportStyles.THIN_BORDER
            det_cell.alignment = ReportStyles.ALIGN_LEFT

            row += 1

        return row

    @staticmethod
    def _compute_health(all_data):
        """Derive traffic-light statuses from collected data."""
        checks = []

        # Cost health
        cost = all_data.get("cost", {})
        pct = cost.get("pct_change", 0)
        if pct >= ReportStyles.COST_INCREASE_CRITICAL:
            checks.append(("Cost Trend", "🔴 CRITICAL", f"Cost increased {pct:.1f}% WoW"))
        elif pct >= ReportStyles.COST_INCREASE_WARNING:
            checks.append(("Cost Trend", "🟡 WARNING", f"Cost increased {pct:.1f}% WoW"))
        else:
            checks.append(("Cost Trend", "🟢 HEALTHY", "Cost within normal range"))

        # CPU health
        ec2_list = all_data.get("ec2", [])
        cpus = [i.get("cpu_avg") for i in ec2_list if i.get("cpu_avg") is not None]
        if cpus:
            max_cpu = max(cpus)
            if max_cpu >= ReportStyles.CPU_CRITICAL:
                checks.append(("CPU Utilisation", "🔴 CRITICAL", f"Peak {max_cpu:.1f}% CPU detected"))
            elif max_cpu >= ReportStyles.CPU_WARNING:
                checks.append(("CPU Utilisation", "🟡 WARNING", f"Peak {max_cpu:.1f}% CPU detected"))
            else:
                checks.append(("CPU Utilisation", "🟢 HEALTHY", f"All instances below {ReportStyles.CPU_WARNING}%"))
        else:
            checks.append(("CPU Utilisation", "🟢 HEALTHY", "No data / all nominal"))

        # Memory health
        mems = [i.get("memory_avg") for i in ec2_list if i.get("memory_avg") is not None]
        if mems:
            max_mem = max(mems)
            if max_mem >= ReportStyles.MEMORY_CRITICAL:
                checks.append(("Memory Utilisation", "🔴 CRITICAL", f"Peak {max_mem:.1f}% memory"))
            elif max_mem >= ReportStyles.MEMORY_WARNING:
                checks.append(("Memory Utilisation", "🟡 WARNING", f"Peak {max_mem:.1f}% memory"))
            else:
                checks.append(("Memory Utilisation", "🟢 HEALTHY", "All instances below threshold"))
        else:
            checks.append(("Memory Utilisation", "🟢 HEALTHY", "No data / all nominal"))

        # WAF health
        waf_list = all_data.get("waf", [])
        total_blocked = sum(w.get("blocked_requests", 0) for w in waf_list)
        total_reqs = sum(w.get("total_requests", 0) for w in waf_list)
        if total_reqs > 0:
            block_pct = (total_blocked / total_reqs) * 100
            if block_pct > 10:
                checks.append(("WAF Security", "🟡 WARNING", f"{block_pct:.1f}% of requests blocked"))
            else:
                checks.append(("WAF Security", "🟢 HEALTHY", f"{total_blocked:,} requests blocked"))
        else:
            checks.append(("WAF Security", "🟢 HEALTHY", "No WAF traffic recorded"))

        # RDS health
        rds_list = all_data.get("rds", [])
        if rds_list:
            rds_cpus = [r.get("cpu_utilization") for r in rds_list if r.get("cpu_utilization") is not None]
            if rds_cpus and max(rds_cpus) >= ReportStyles.CPU_CRITICAL:
                checks.append(("RDS Health", "🔴 CRITICAL", f"Peak CPU {max(rds_cpus):.1f}%"))
            elif rds_cpus and max(rds_cpus) >= ReportStyles.CPU_WARNING:
                checks.append(("RDS Health", "🟡 WARNING", f"Peak CPU {max(rds_cpus):.1f}%"))
            else:
                checks.append(("RDS Health", "🟢 HEALTHY", "All databases healthy"))
        else:
            checks.append(("RDS Health", "🟢 HEALTHY", "No RDS instances in scope"))

        # Security Findings health (Amazon Inspector + GuardDuty)
        inspector = all_data.get("inspector") or {}
        guardduty = all_data.get("guardduty") or {}
        insp_counts = inspector.get("severity_counts", {}) or {}
        gd_counts = guardduty.get("severity_counts", {}) or {}

        insp_status = inspector.get("collection_status", "ok")
        gd_status = guardduty.get("collection_status", "ok")

        crit = int(insp_counts.get("CRITICAL", 0) or 0)
        high = int(insp_counts.get("HIGH", 0) or 0) + int(gd_counts.get("HIGH", 0) or 0)
        medium = int(insp_counts.get("MEDIUM", 0) or 0) + int(
            gd_counts.get("MEDIUM", 0) or 0
        )

        if "error" in (insp_status, gd_status):
            checks.append((
                "Security Findings",
                "🟡 WARNING",
                "Collection failed / access denied — results incomplete",
            ))
        elif crit > 0 or high > 0:
            checks.append((
                "Security Findings",
                "🔴 CRITICAL",
                f"{crit} Critical, {high} High finding(s) this week",
            ))
        elif medium > 0:
            checks.append((
                "Security Findings",
                "🟡 WARNING",
                f"{medium} Medium finding(s) this week",
            ))
        else:
            checks.append((
                "Security Findings",
                "🟢 HEALTHY",
                "No high-severity Inspector/GuardDuty findings",
            ))

        return checks
