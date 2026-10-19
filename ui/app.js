/* 개인정보 가리개 - 화면 동작 (모양 확인용).
   기능은 연결되어 있지 않다: 모든 버튼은 "아직 연결 안 됨" 알림창만 띄운다.
   외부 주소는 쓰지 않는다. */
(function () {
  "use strict";

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  /* ---------------------------------------------------------- 알림창 */
  const dialogScrim = $("#dialog-scrim");
  const dialogTitle = $("#dialog-title");
  const dialogOk = $("#dialog-ok");
  const dialogCancel = $("#dialog-cancel");
  let dialogReturnFocus = null;

  function showAlert(title) {
    dialogTitle.textContent = title;
    dialogReturnFocus = document.activeElement;
    dialogScrim.hidden = false;
    requestAnimationFrame(() => dialogScrim.classList.add("is-open"));
    dialogOk.focus();
  }

  function closeAlert() {
    dialogScrim.classList.remove("is-open");
    setTimeout(() => { dialogScrim.hidden = true; }, 200);
    if (dialogReturnFocus && dialogReturnFocus.focus) dialogReturnFocus.focus();
  }

  dialogOk.addEventListener("click", closeAlert);
  dialogCancel.addEventListener("click", closeAlert);
  dialogScrim.addEventListener("click", (e) => { if (e.target === dialogScrim) closeAlert(); });

  // 알림창 안에서 Tab 이 밖으로 나가지 않게
  dialogScrim.addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    const first = dialogCancel, last = dialogOk;
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  });

  // 기능이 아직 없는 모든 버튼
  document.addEventListener("click", (e) => {
    const stub = e.target.closest("[data-stub]");
    if (stub && !stub.disabled) showAlert("아직 연결 안 됨");
  });

  /* ---------------------------------------------------------- 토스트 */
  const toast = $("#toast");
  let toastTimer = null;
  function showToast(message) {
    toast.textContent = message;
    toast.classList.add("is-show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toast.classList.remove("is-show"), 1500);
  }

  /* ---------------------------------------------------------- 펼침 / 접힘 (900px 기준) */
  const panels = $("#panels");
  const paneMask = $("#pane-mask");
  const paneRestore = $("#pane-restore");
  const foldTabs = $$("[data-fold-tab]");
  const foldQuery = window.matchMedia("(max-width: 899px)");

  function syncFold() {
    const folded = foldQuery.matches;
    const active = panels.dataset.active;
    // 접힌 상태에서는 보이지 않는 패널을 키보드·스크린리더에서 뺀다
    paneMask.inert = folded && active !== "mask";
    paneRestore.inert = folded && active !== "restore";
    foldTabs.forEach((t) => t.setAttribute("aria-selected", String(t.dataset.foldTab === active)));
  }

  function setActive(which) {
    if (panels.dataset.active === which) return;
    // 탭을 바꿀 때는 너비 애니메이션 없이 바로 바꾼다 (펼침↔접힘 전환에서만 0.4초 애니메이션)
    panels.classList.add("is-instant");
    panels.dataset.active = which;
    syncFold();
    requestAnimationFrame(() => requestAnimationFrame(() => panels.classList.remove("is-instant")));
  }

  foldTabs.forEach((t) => t.addEventListener("click", () => setActive(t.dataset.foldTab)));
  foldQuery.addEventListener("change", syncFold);
  syncFold();

  /* ---------------------------------------------------------- 세그먼트 (라디오처럼 하나만 선택) */
  function bindSegmented(group, onChange) {
    const items = $$('[role="radio"]', group);
    items.forEach((item) => {
      item.addEventListener("click", () => {
        items.forEach((i) => i.setAttribute("aria-checked", String(i === item)));
        if (onChange) onChange(item);
      });
    });
    group.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowRight" && e.key !== "ArrowLeft") return;
      const idx = items.findIndex((i) => i.getAttribute("aria-checked") === "true");
      const next = items[(idx + (e.key === "ArrowRight" ? 1 : items.length - 1)) % items.length];
      next.click();
      next.focus();
    });
  }

  bindSegmented($("#mode-seg"));

  /* ---------------------------------------------------------- 파일 끌어다 놓기 (파일 이름만 보여 줌) */
  const dropCard = $("#drop-card");
  const dropEmpty = $("#drop-empty");
  const dropPicked = $("#drop-picked");
  const KINDS = {
    pdf: "PDF 문서", docx: "Word 문서", xlsx: "엑셀 통합 문서", xls: "엑셀 문서",
    pptx: "PowerPoint 문서", csv: "CSV 표", txt: "텍스트", md: "마크다운",
    hwp: "한글 문서", hwpx: "한글 문서",
  };

  // 창 어디에 놓아도 브라우저가 파일을 열어 버리지 않게 막는다
  ["dragover", "drop"].forEach((type) => window.addEventListener(type, (e) => e.preventDefault()));

  let dragDepth = 0;
  dropCard.addEventListener("dragenter", (e) => { e.preventDefault(); dragDepth++; dropCard.classList.add("is-dragover"); });
  dropCard.addEventListener("dragleave", () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) dropCard.classList.remove("is-dragover"); });
  dropCard.addEventListener("drop", (e) => {
    e.preventDefault();
    dragDepth = 0;
    dropCard.classList.remove("is-dragover");
    const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (!file) return;
    const ext = (file.name.split(".").pop() || "").toLowerCase();
    $("#file-name").textContent = file.name;
    $("#file-kind").textContent = KINDS[ext] || (ext ? ext.toUpperCase() + " 파일" : "파일");
    dropEmpty.hidden = true;
    dropPicked.hidden = false;
  });

  /* ---------------------------------------------------------- 미리보기: 코드 칩 강조 */
  const CODE = /\[(?:학생전화|보호자전화|주소|생년월일)\d{3,}(?:또는\d{3,})*\]|\[(?:전화번호|주민번호|이메일)\]|(?:학생|보호자)\d{3,}(?:또는\d{3,})*|S\d{3,}(?:또는\d{3,})*/g;

  function renderPreview(el, text) {
    el.textContent = "";
    let last = 0;
    for (const m of text.matchAll(CODE)) {
      if (m.index > last) el.append(document.createTextNode(text.slice(last, m.index)));
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = m[0];
      el.append(chip);
      last = m.index + m[0].length;
    }
    if (last < text.length) el.append(document.createTextNode(text.slice(last)));
  }

  // 확인용 예시 문장
  renderPreview($("#preview"), "학생003은 S003이며 연락처는 [학생전화003]입니다. 보호자003과 상담함.");

  /* ---------------------------------------------------------- 시트 (확인 필요 목록) */
  const sheet = $("#sheet");
  const sheetScrim = $("#sheet-scrim");
  const openSheetLink = $("#open-sheet");

  function openSheet() {
    sheet.inert = false;
    sheetScrim.hidden = false;
    requestAnimationFrame(() => { sheet.classList.add("is-open"); sheetScrim.classList.add("is-open"); });
    $("#close-sheet").focus();
  }

  function closeSheet() {
    sheet.classList.remove("is-open");
    sheetScrim.classList.remove("is-open");
    sheet.inert = true;
    setTimeout(() => { if (!sheet.classList.contains("is-open")) sheetScrim.hidden = true; }, 400);
    openSheetLink.focus();
  }

  openSheetLink.addEventListener("click", openSheet);
  $("#close-sheet").addEventListener("click", closeSheet);
  $("#sheet-handle").addEventListener("click", closeSheet);
  sheetScrim.addEventListener("click", closeSheet);

  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    if (!dialogScrim.hidden) closeAlert();
    else if (sheet.classList.contains("is-open")) closeSheet();
  });

  /* ---------------------------------------------------------- 결과 상태 표시 */
  const resultCard = $("#result-card");
  const statusText = $("#status-text");
  const copyResult = $("#copy-result");

  const STATES = {
    safe:   { text: "남은 개인정보 없음", list: false, copy: true },
    warn:   { text: "확인 필요 3건",      list: true,  copy: true },
    danger: { text: "남은 개인정보 발견", list: false, copy: false },  // 위험: 결과 복사 비활성
  };

  function setResultState(name) {
    const s = STATES[name];
    resultCard.dataset.state = name;
    statusText.textContent = s.text;
    openSheetLink.hidden = !s.list;
    copyResult.disabled = !s.copy;
    if (!s.list && sheet.classList.contains("is-open")) closeSheet();
  }

  /* ---------------------------------------------------------- ▼ 임시 미리보기 스위치 (나중에 지울 것) ▼ */
  const STATE_NAMES = { safe: "안전", warn: "확인 필요", danger: "위험" };
  bindSegmented($("#temp-state"), (item) => {
    setResultState(item.dataset.tempState);
    showToast("결과 상태: " + STATE_NAMES[item.dataset.tempState]);
  });
  bindSegmented($("#temp-issues"), (item) => {
    $("#issues-card").hidden = item.dataset.tempIssues === "hide";
  });
  setResultState("safe");
  /* ---------------------------------------------------------- ▲ 임시 미리보기 스위치 끝 ▲ */
})();
