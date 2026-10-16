(() => {
  const $ = (selector) => document.querySelector(selector);
  const state = { reports: {}, report: null, view: "HN" };
  const sheetLabels = {
    dinner: "Check-in tiệc tối",
    dinner_invited: "Khách mời tiệc tối",
    afternoon: "Check-in buổi chiều",
    registered: "Đăng ký tham dự",
  };
  const elements = {
    file: $("#excel-file"), choose: $("#choose-file"), emptyChoose: $("#empty-choose"),
    notice: $("#notice"), empty: $("#empty-state"), dashboard: $("#dashboard"),
    compareDashboard: $("#compare-dashboard"), uploadRegion: $("#upload-region"),
    chartCategory: $("#chart-category"), vipFilter: $("#vip-filter"), inviterFilter: $("#inviter-filter"),
    chart: $("#chart"), chartEmpty: $("#chart-empty"), chartTotal: $("#chart-total"),
    chartAttribution: $("#chart-attribution"),
    employeeRows: $("#employee-rows"), employeeSearch: $("#employee-search"),
    dialog: $("#detail-dialog"),
  };

  function notice(message, isError = true) {
    elements.notice.textContent = message;
    elements.notice.hidden = !message;
    elements.notice.style.background = isError ? "#fff3e8" : "#eaf5e8";
    elements.notice.style.borderColor = isError ? "#f1cfb1" : "#c6e5c4";
    elements.notice.style.color = isError ? "#91501d" : "#2f7135";
  }

  async function uploadFile(files) {
    if (!files || files.length === 0) return;
    const region = elements.uploadRegion.value;
    const names = [...files].map((file) => file.name);
    $("#upload-name").textContent = names.join(", ");
    notice(`Đang đọc ${files.length} file Excel cho ${region === "HN" ? "Hà Nội" : "TP. Hồ Chí Minh"}…`, false);
    const form = new FormData();
    for (const file of files) form.append("files", file);
    form.append("region", region);
    try {
      const response = await fetch("/api/upload", { method: "POST", body: form });
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.message || "Không thể xử lý các file Excel.");
      state.reports = result.reports || { [region]: result.report };
      state.view = region;
      elements.uploadRegion.value = region;
      notice(`Đã cập nhật báo cáo ${region === "HN" ? "Hà Nội" : "TP. Hồ Chí Minh"} từ ${files.length} file Excel.`, false);
      render();
    } catch (error) {
      notice(error.message || "Có lỗi khi tải file.");
    }
  }

  function render() {
    const isComparison = state.view === "compare";
    const report = state.reports[state.view] || null;
    state.report = report;
    elements.compareDashboard.hidden = !isComparison;
    elements.dashboard.hidden = isComparison || !report;
    elements.empty.hidden = isComparison || Boolean(report);
    $("#region-eyebrow").textContent = isComparison
      ? "BÁO CÁO · SO SÁNH HAI MIỀN"
      : `DASHBOARD · ${state.view === "HN" ? "HÀ NỘI" : "TP. HỒ CHÍ MINH"}`;
    $("#empty-message").textContent = `Chưa có báo cáo ${state.view === "HN" ? "Hà Nội" : "TP. Hồ Chí Minh"}. Chọn miền cần cập nhật rồi tải lên 4 file Excel riêng: đăng ký, check-in buổi chiều, khách mời tiệc tối và chỉ tiêu nhân viên.`;
    document.querySelectorAll(".region-tab").forEach((button) => {
      const active = button.dataset.view === state.view;
      button.classList.toggle("active", active);
      button.setAttribute("aria-pressed", String(active));
    });
    if (state.view !== "compare") elements.uploadRegion.value = state.view;
    if (isComparison) {
      renderComparison();
      return;
    }
    if (!report) return;
    $("#metric-employees").textContent = report.totals.employees.toLocaleString("vi-VN");
    $("#metric-registered").textContent = report.totals.registered.toLocaleString("vi-VN");
    $("#metric-afternoon").textContent = report.totals.afternoon.toLocaleString("vi-VN");
    $("#metric-dinner-invited").textContent = (report.totals.dinner_invited ?? report.totals.dinner).toLocaleString("vi-VN");
    $("#metric-dinner").textContent = report.totals.dinner.toLocaleString("vi-VN");
    renderSourceSummary();
    renderWarnings();
    renderInviters();
    renderChart();
    renderEmployees();
  }

  function renderComparison() {
    const hanoi = state.reports.HN;
    const hcm = state.reports.HCM;
    const available = Boolean(hanoi && hcm);
    $("#comparison-empty").hidden = available;
    $("#comparison-table-wrap").hidden = !available;
    const body = $("#comparison-rows");
    body.replaceChildren();
    if (!available) return;

    const dinnerInvited = (report) => report.totals.dinner_invited ?? report.totals.dinner;
    const metrics = [
      ["Nhân viên trong danh sách target", hanoi.totals.employees, hcm.totals.employees],
      ["Đăng ký tham dự", hanoi.totals.registered, hcm.totals.registered],
      ["Check-in buổi chiều", hanoi.totals.afternoon, hcm.totals.afternoon],
      ["Khách mời tiệc tối", dinnerInvited(hanoi), dinnerInvited(hcm)],
      ["Check-in tiệc tối", hanoi.totals.dinner, hcm.totals.dinner],
      [
        "Tỷ lệ check-in tiệc tối",
        dinnerInvited(hanoi) ? hanoi.totals.dinner / dinnerInvited(hanoi) : null,
        dinnerInvited(hcm) ? hcm.totals.dinner / dinnerInvited(hcm) : null,
        true,
      ],
    ];
    for (const [label, hanoiValue, hcmValue, isRate = false] of metrics) {
      const row = document.createElement("tr");
      const values = [
        label,
        formatComparisonValue(hanoiValue, isRate),
        formatComparisonValue(hcmValue, isRate),
        formatComparisonDifference(hanoiValue, hcmValue, isRate),
      ];
      for (const value of values) {
        const cell = document.createElement("td");
        cell.textContent = value;
        row.append(cell);
      }
      body.append(row);
    }
  }

  function formatComparisonValue(value, isRate) {
    if (value == null || !Number.isFinite(value)) return "—";
    return isRate
      ? new Intl.NumberFormat("vi-VN", { style: "percent", maximumFractionDigits: 1 }).format(value)
      : value.toLocaleString("vi-VN");
  }

  function formatComparisonDifference(hanoiValue, hcmValue, isRate) {
    if (hanoiValue == null || hcmValue == null) return "—";
    const difference = hcmValue - hanoiValue;
    if (isRate) {
      const points = difference * 100;
      return `${points > 0 ? "+" : ""}${new Intl.NumberFormat("vi-VN", { maximumFractionDigits: 1 }).format(points)} điểm %`;
    }
    return `${difference > 0 ? "+" : ""}${difference.toLocaleString("vi-VN")}`;
  }

  function renderSourceSummary() {
    const container = $("#source-summary");
    container.replaceChildren();
    if (!state.report.sources) {
      container.hidden = true;
      return;
    }
    container.hidden = false;
    const labels = {
      registered: "Đăng ký",
      afternoon: "Check-in chiều",
      dinner: "Khách mời tối",
      employees: "Chỉ tiêu nhân viên",
    };
    const heading = document.createElement("strong");
    heading.textContent = "Đã đọc đủ nguồn dữ liệu:";
    container.append(heading);
    for (const source of state.report.sources) {
      const item = document.createElement("span");
      const guestCount = source.category === "dinner"
        ? ` · ${source.guests.toLocaleString("vi-VN")} khách có số lượng${source.unquantified_records ? ` · ${source.unquantified_records.toLocaleString("vi-VN")} dòng thiếu SL (${source.unquantified_checkins.toLocaleString("vi-VN")} có dấu check-in)` : ""}`
        : "";
      item.textContent = `${labels[source.category] || source.category}: ${source.records.toLocaleString("vi-VN")} dòng${guestCount}`;
      item.title = source.filename;
      container.append(item);
    }
  }

  function renderWarnings() {
    const container = $("#report-warnings");
    container.replaceChildren();
    const warnings = state.report.warnings || [];
    container.hidden = warnings.length === 0;
    for (const warning of warnings) {
      const item = document.createElement("p");
      item.textContent = warning;
      container.append(item);
    }
  }

  function renderInviters() {
    const category = elements.chartCategory.value;
    const groupLabel = state.report.chart_group_by?.[category] || "Người mời";
    const previous = elements.inviterFilter.value || "all";
    const inviters = new Set();
    (state.report.charts[category] || []).forEach((item) => {
      if (item.name && item.name !== "Chưa xác định") inviters.add(item.name);
    });
    $("#chart-group-label").textContent = groupLabel;
    const allLabel = groupLabel === "Kênh" ? "Tất cả kênh" : "Tất cả nhân viên";
    elements.inviterFilter.replaceChildren(new Option(allLabel, "all"));
    [...inviters].sort((a, b) => a.localeCompare(b, "vi")).forEach((name) => elements.inviterFilter.add(new Option(name, name)));
    if ([...elements.inviterFilter.options].some((option) => option.value === previous)) elements.inviterFilter.value = previous;
  }

  function renderChart() {
    const category = elements.chartCategory.value;
    const groupLabel = state.report.chart_group_by?.[category] || "Người mời";
    const filter = elements.vipFilter.value;
    const selectedInviter = elements.inviterFilter.value;
    $("#chart-title").textContent = groupLabel === "Kênh"
      ? `${sheetLabels[category]} theo kênh`
      : sheetLabels[category];
    const rows = (state.report.charts[category] || [])
      .filter((item) => selectedInviter === "all" || item.name === selectedInviter)
      .map((item) => ({
        ...item,
        vip: filter === "non_vip" || filter === "unclassified" ? 0 : item.vip,
        non_vip: filter === "vip" || filter === "unclassified" ? 0 : item.non_vip,
        unclassified: filter === "vip" || filter === "non_vip" ? 0 : (item.unclassified || 0),
      }))
      .map((item) => ({ ...item, visibleCount: item.vip + item.non_vip + item.unclassified }))
      .sort((a, b) => b.visibleCount - a.visibleCount || a.name.localeCompare(b.name, "vi"));
    const max = Math.max(1, ...rows.map((item) => item.visibleCount));
    let total = 0;
    let unattributed = 0;
    if (!elements.chartAttribution) {
      elements.chartAttribution = document.createElement("p");
      elements.chartAttribution.id = "chart-attribution";
      elements.chartAttribution.className = "chart-attribution";
      elements.chart.insertAdjacentElement("afterend", elements.chartAttribution);
    }
    elements.chart.replaceChildren();
    for (const item of rows) {
      const vip = item.vip;
      const regular = item.non_vip;
      const unclassified = item.unclassified;
      const count = item.visibleCount;
      total += count;
      if (item.name === "Chưa xác định") unattributed += count;

      const row = document.createElement("div");
      row.className = "chart-row";
      row.title = `${item.name}: ${vip} VIP, ${regular} không VIP, ${unclassified} chưa phân loại`;
      const label = document.createElement("span");
      label.className = "chart-row-label";
      label.textContent = item.name;
      const track = document.createElement("div");
      track.className = "chart-track";
      const bar = document.createElement("div");
      bar.className = "chart-bar";
      bar.style.width = `${(count / max) * 100}%`;
      for (const [className, value] of [["vip", vip], ["normal", regular], ["unclassified", unclassified]]) {
        if (!value) continue;
        const segment = document.createElement("span");
        segment.className = `chart-segment ${className}`;
        segment.style.width = `${(value / count) * 100}%`;
        if (value / count >= 0.16) segment.textContent = value.toLocaleString("vi-VN");
        bar.append(segment);
      }
      track.append(bar);
      const number = document.createElement("strong");
      number.className = "chart-row-total";
      number.textContent = count.toLocaleString("vi-VN");
      row.append(label, track, number);
      elements.chart.append(row);
    }
    elements.chartTotal.textContent = `${total.toLocaleString("vi-VN")} khách · ${rows.length} nhóm`;
    elements.chart.hidden = rows.length === 0;
    elements.chartEmpty.hidden = rows.length !== 0;
    elements.chartAttribution.hidden = unattributed === 0;
    elements.chartAttribution.textContent = unattributed
      ? `${unattributed.toLocaleString("vi-VN")} khách chưa có thông tin ${groupLabel.toLocaleLowerCase("vi")}; hãy kiểm tra dữ liệu nguồn.`
      : "";
    elements.chart.setAttribute("aria-label", `Biểu đồ ${sheetLabels[category]}, ${total} khách`);
  }

  function renderEmployees() {
    const query = $("#employee-search").value.trim().toLocaleLowerCase("vi");
    elements.employeeRows.replaceChildren();
    const employees = state.report.employees.filter((employee) => employee.name.toLocaleLowerCase("vi").includes(query));
    let previousDepartment = null;
    for (const employee of employees) {
      const row = document.createElement("tr");
      const department = (employee.department || "").trim().toLocaleLowerCase("vi");
      if (department !== previousDepartment) row.classList.add("department-start");
      previousDepartment = department;
      const nameCell = document.createElement("td");
      const link = document.createElement("button");
      link.className = "employee-link";
      link.type = "button";
      link.textContent = employee.name;
      link.addEventListener("click", () => showDetails(employee));
      nameCell.append(link);
      row.append(nameCell);
      for (const key of ["position", "department", "target", "invited", "afternoon", "dinner_invited", "dinner"]) {
        const cell = document.createElement("td");
        const value = key === "dinner_invited"
          ? (employee.dinner_invited ?? employee.dinner)
          : employee[key];
        cell.textContent = value === "" || value == null ? "—" : String(value);
        row.append(cell);
      }
      const attendanceCell = document.createElement("td");
      attendanceCell.className = "attendance-rate";
      attendanceCell.textContent = employee.attendance_rate == null
        ? "—"
        : new Intl.NumberFormat("vi-VN", { style: "percent", maximumFractionDigits: 1 }).format(employee.attendance_rate);
      row.append(attendanceCell);
      elements.employeeRows.append(row);
    }
    const unmatched = state.report.unmatched;
    const unmatchedValues = unmatched && [
      unmatched.registered,
      unmatched.afternoon,
      unmatched.dinner_invited,
      unmatched.dinner,
    ];
    const hasUnmatchedSummary = unmatchedValues?.some((value) => value > 0)
      && "chưa đối chiếu target".includes(query);
    if (hasUnmatchedSummary) {
      const row = document.createElement("tr");
      row.className = "unmatched-summary";
      const label = document.createElement("td");
      label.textContent = "CHƯA ĐỐI CHIẾU TARGET";
      row.append(label);
      for (const value of [
        "Người mời không khớp",
        "Chưa đối chiếu",
        "—",
        unmatched.registered,
        unmatched.afternoon,
        unmatched.dinner_invited,
        unmatched.dinner,
      ]) {
        const cell = document.createElement("td");
        cell.textContent = String(value);
        row.append(cell);
      }
      const attendanceCell = document.createElement("td");
      attendanceCell.textContent = "—";
      row.append(attendanceCell);
      elements.employeeRows.append(row);
    }
    if (employees.length === 0 && !hasUnmatchedSummary) {
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      cell.colSpan = 9;
      cell.textContent = "Không tìm thấy nhân viên phù hợp.";
      row.append(cell);
      elements.employeeRows.append(row);
    }
  }

  function showDetails(employee) {
    $("#detail-name").textContent = employee.name;
    $("#detail-meta").textContent = [employee.position, employee.department].filter(Boolean).join(" · ") || "Thông tin nhân viên";
    const summary = $("#detail-summary");
    summary.replaceChildren();
    for (const [label, key] of [
      ["Đăng ký", "invited"],
      ["Check-in chiều", "afternoon"],
      ["Mời tiệc tối", "dinner_invited"],
      ["Check-in tiệc tối", "dinner"],
    ]) {
      const pill = document.createElement("span");
      pill.className = "summary-pill";
      const value = key === "dinner_invited" ? (employee.dinner_invited ?? employee.dinner) : employee[key];
      pill.textContent = `${label}: ${value}`;
      summary.append(pill);
    }
    const body = $("#guest-rows");
    body.replaceChildren();
    const guests = [...employee.guests].sort((a, b) => (a.name || "").localeCompare(b.name || "", "vi"));
    if (!guests.length) {
      const row = document.createElement("tr");
      const cell = document.createElement("td");
      cell.colSpan = 11;
      cell.textContent = "Chưa có khách được liên kết với nhân viên này.";
      row.append(cell);
      body.append(row);
    }
    for (const guest of guests) {
      const row = document.createElement("tr");
      const name = document.createElement("td");
      name.textContent = guest.name || "Chưa có tên";
      row.append(name);
      for (const value of [guest.customer_id, guest.phone]) {
        const cell = document.createElement("td");
        cell.textContent = value || "—";
        row.append(cell);
      }
      const sourceCell = document.createElement("td");
      sourceCell.textContent = ({
        registered: "Đăng ký",
        afternoon: "Check-in chiều",
        dinner: "Mời tối",
      })[guest.source] || "Workbook";
      row.append(sourceCell);
      const quantityCell = document.createElement("td");
      quantityCell.textContent = guest.quantity == null ? "Chưa rõ" : String(guest.quantity);
      row.append(quantityCell);
      const vipCell = document.createElement("td");
      const vipTag = document.createElement("span");
      vipTag.className = guest.vip ? "vip-tag" : "normal-tag";
      vipTag.textContent = guest.vip
        ? "VIP"
        : guest.vip_classified === false ? "Chưa rõ" : "Không VIP";
      vipCell.append(vipTag);
      row.append(vipCell);
      for (const status of ["registered", "afternoon", "dinner_invited", "dinner"]) {
        const cell = document.createElement("td");
        const mark = document.createElement("span");
        mark.className = `check-mark${guest[status] ? " checked" : ""}`;
        mark.textContent = guest[status] ? "✓" : "–";
        mark.setAttribute("aria-label", guest[status] ? "Đã ghi nhận" : "Chưa ghi nhận");
        cell.append(mark);
        row.append(cell);
      }
      const rawCell = document.createElement("td");
      if (guest.source_data && Object.keys(guest.source_data).length) {
        const details = document.createElement("details");
        const summary = document.createElement("summary");
        summary.textContent = "Xem";
        const list = document.createElement("dl");
        for (const [label, value] of Object.entries(guest.source_data)) {
          const term = document.createElement("dt");
          term.textContent = label;
          const description = document.createElement("dd");
          description.textContent = value || "—";
          list.append(term, description);
        }
        details.append(summary, list);
        rawCell.append(details);
      } else {
        rawCell.textContent = "—";
      }
      row.append(rawCell);
      body.append(row);
    }
    elements.dialog.showModal();
  }

  elements.choose.addEventListener("click", () => elements.file.click());
  elements.emptyChoose.addEventListener("click", () => elements.file.click());
  elements.file.addEventListener("change", () => uploadFile(elements.file.files));
  document.querySelectorAll(".region-tab").forEach((button) => {
    button.addEventListener("click", () => {
      state.view = button.dataset.view;
      render();
    });
  });
  elements.chartCategory.addEventListener("change", () => {
    renderInviters();
    renderChart();
  });
  elements.vipFilter.addEventListener("change", renderChart);
  elements.inviterFilter.addEventListener("change", renderChart);
  elements.employeeSearch.addEventListener("input", renderEmployees);
  $("#close-detail").addEventListener("click", () => elements.dialog.close());
  elements.dialog.addEventListener("click", (event) => {
    if (event.target === elements.dialog) elements.dialog.close();
  });
  fetch("/api/report").then((response) => response.ok ? response.json() : null).then((result) => {
    if (result?.ok) {
      state.reports = result.reports || (result.report ? { HN: result.report } : {});
      if (!state.reports.HN && state.reports.HCM) state.view = "HCM";
      elements.uploadRegion.value = state.view === "compare" ? "HN" : state.view;
      notice("Đã khôi phục báo cáo đã lưu.", false);
      render();
    }
  }).catch(() => notice("Không thể kết nối để tải báo cáo đã lưu.", true));
})();
