import React, { useEffect, useMemo, useState } from "react";
import { motion } from "framer-motion";

/**
 * نظام متابعة مشاريع المساجد (واجهة تجريبية)
 * - 4 مراحل = 25% لكل مرحلة
 * - العداد (بالأيام) يُحسب للمرحلة الحالية فقط
 * - إذا انتهت مدة المرحلة ولم تُعتمد: يصبح العداد سالب ويستمر حتى اعتماد المرحلة
 * - عند اعتماد المرحلة: تُسجل تاريخ الاعتماد وتبدأ المرحلة التالية من تاريخ الاعتماد
 *
 * ملاحظة:
 * - هذا ملف واجهة واحد (React) مع بيانات تجريبية. اربطه لاحقًا بقاعدة بيانات/API.
 */

// -----------------------------
// Helpers
// -----------------------------

const MS_PER_DAY = 24 * 60 * 60 * 1000;

function startOfDay(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function daysBetween(a, b) {
  // b - a (بالأيام)
  const aa = startOfDay(a).getTime();
  const bb = startOfDay(b).getTime();
  return Math.round((bb - aa) / MS_PER_DAY);
}

function formatDurationDaysToMonthsDays(totalDays) {
  // تقريب بسيط: الشهر = 30 يوم
  const m = Math.floor(totalDays / 30);
  const d = Math.abs(totalDays % 30);
  return `${m} شهر + ${d} يوم`;
}

function clampInt(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function stageLabel(i) {
  return `المرحلة ${i}`;
}

function statusPill(status, subtitle) {
  const isDone = status === "تم";
  const base =
    "inline-flex items-center justify-center px-4 py-2 rounded-full text-sm font-semibold border";
  const cls = isDone
    ? "bg-white border-emerald-200 text-emerald-700"
    : "bg-rose-50 border-rose-100 text-rose-700";
  return (
    <div className={`${base} ${cls}`}>
      <div className="leading-tight text-center">
        <div>{status}</div>
        {!isDone && subtitle ? (
          <div className="text-[11px] opacity-70">{subtitle}</div>
        ) : null}
      </div>
    </div>
  );
}

function computeStageLengthDays(durationDays) {
  // سياسة التقسيم:
  // - 3 مراحل أولى = floor(25%)
  // - المرحلة الرابعة = المتبقي لضمان المجموع = مدة المشروع
  const q = durationDays * 0.25;
  const s = Math.floor(q);
  const s1 = s;
  const s2 = s;
  const s3 = s;
  const s4 = Math.max(1, durationDays - (s1 + s2 + s3));
  return [s1, s2, s3, s4];
}

function getCurrentStage(project) {
  // المرحلة الحالية = أول مرحلة ليست "تم"
  for (let i = 1; i <= 4; i++) {
    if (project.stages[i - 1].status !== "تم") return i;
  }
  return 4; // مكتملة
}

function getStageStartDate(project, stageIndex /* 1..4 */) {
  // المرحلة 1: start_date
  // المرحلة 2: stage1_approved_at
  // المرحلة 3: stage2_approved_at
  // المرحلة 4: stage3_approved_at
  if (stageIndex === 1) return project.start_date;
  const prevApprovedAt = project.stages[stageIndex - 2].approved_at;
  return prevApprovedAt || project.start_date;
}

function computeCounterDays(project, today = new Date()) {
  const cur = getCurrentStage(project);

  // إذا مكتمل (كلها تم): العداد = 0
  const allDone = project.stages.every((s) => s.status === "تم");
  if (allDone) return 0;

  const stageLengths = computeStageLengthDays(project.duration_days);
  const stageLen = stageLengths[cur - 1];
  const stageStart = getStageStartDate(project, cur);
  const stageEnd = new Date(startOfDay(stageStart).getTime() + stageLen * MS_PER_DAY);

  // counter = stage_end - today (متبقي موجب، متأخر سالب)
  return daysBetween(today, stageEnd);
}

function computeStatusSummary(projects) {
  const total = projects.length;
  const completed = projects.filter((p) => p.stages.every((s) => s.status === "تم")).length;
  const late = projects.filter((p) => computeCounterDays(p) < 0).length;
  return { total, completed, late };
}

// -----------------------------
// Demo Data (مستوردة من ملف Word)
// -----------------------------

const importedRawProjects = [];

function normalizeStages(stage_statuses) {
  // ضمان التسلسل: إذا مرحلة ما "مطلوب"، اجعل كل ما بعدها "مطلوب"
  const out = [...stage_statuses].map((s) => (s && s.trim() ? s.trim() : "مطلوب"));
  let blocked = false;
  for (let i = 0; i < out.length; i++) {
    if (blocked) out[i] = "مطلوب";
    if (out[i] !== "تم") blocked = true;
  }
  // طول 4 دائمًا
  while (out.length < 4) out.push("مطلوب");
  return out.slice(0, 4);
}

function hydrateImportedProjects(raw, today = new Date()) {
  const t0 = startOfDay(today);
  return raw.map((r) => {
    const stage_statuses = normalizeStages(r.stage_statuses);
    const stageLens = computeStageLengthDays(r.duration_days);

    // نبني تواريخ اعتماد افتراضية متسلسلة (لتمكين منطق العداد).
    // لعدم توفر تواريخ البداية من ملف Word، نعتبر أن المرحلة الحالية بدأت اليوم.
    const firstNotDone = stage_statuses.findIndex((s) => s !== "تم");
    const currentStage = firstNotDone === -1 ? 4 : firstNotDone + 1;

    const completedSum = stageLens.slice(0, currentStage - 1).reduce((a, b) => a + b, 0);
    const start_date = new Date(t0.getTime() - completedSum * MS_PER_DAY);

    const stages = stage_statuses.map((st, idx) => {
      if (st !== "تم") return { status: "مطلوب", approved_at: null };
      const approvedAt = new Date(startOfDay(start_date).getTime() + stageLens.slice(0, idx + 1).reduce((a, b) => a + b, 0) * MS_PER_DAY);
      return { status: "تم", approved_at: approvedAt };
    });

    return {
      id: `p-${r.project_no}`,
      project_no: r.project_no,
      contractor_name: r.contractor_name,
      country: r.country,
      duration_days: r.duration_days,
      duration_text: r.duration_text,
      start_date,
      stages,
    };
  });
}

// إذا عندك ملف JSON كامل: استبدل importedRawProjects أو اربطه بـ API
const seedProjects = hydrateImportedProjects(importedRawProjects);

async function loadProjectsFromJson(setProjects) {
  try {
    const res = await fetch("/projects.json", { cache: "no-store" });
    if (!res.ok) return;
    const raw = await res.json();
    setProjects(hydrateImportedProjects(raw));
  } catch {
    // ignore
  }
}

// -----------------------------
// UI
// -----------------------------

function CardStat({ title, value, accent }) {
  const border =
    accent === "green"
      ? "border-emerald-400"
      : accent === "red"
      ? "border-rose-400"
      : "border-blue-400";

  return (
    <div className={`bg-white rounded-2xl shadow-sm border-t-4 ${border} p-6`}>
      <div className="text-4xl font-extrabold text-slate-900 text-center">{value}</div>
      <div className="text-slate-500 text-center mt-2">{title}</div>
    </div>
  );
}

function Select({ value, onChange, children, ariaLabel }) {
  return (
    <select
      aria-label={ariaLabel}
      className="w-full md:w-52 rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
      value={value}
      onChange={(e) => onChange(e.target.value)}
    >
      {children}
    </select>
  );
}

function Input({ value, onChange, placeholder }) {
  return (
    <input
      className="w-full md:w-72 rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
      value={value}
      onChange={(e) => onChange(e.target.value)}
      placeholder={placeholder}
    />
  );
}

function Button({ children, onClick, variant = "primary" }) {
  const cls =
    variant === "primary"
      ? "bg-blue-600 text-white hover:bg-blue-700"
      : variant === "ghost"
      ? "bg-white text-slate-800 hover:bg-slate-50 border border-slate-200"
      : "bg-rose-600 text-white hover:bg-rose-700";
  return (
    <button
      onClick={onClick}
      className={`rounded-xl px-4 py-3 font-semibold shadow-sm transition ${cls}`}
      type="button"
    >
      {children}
    </button>
  );
}

function Modal({ open, title, onClose, children }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center">
      <div className="absolute inset-0 bg-black/40" onClick={onClose} />
      <div className="relative w-[95vw] max-w-2xl rounded-2xl bg-white shadow-xl border border-slate-100">
        <div className="p-5 border-b border-slate-100 flex items-center justify-between">
          <div className="text-lg font-extrabold text-slate-900">{title}</div>
          <button
            onClick={onClose}
            className="rounded-xl px-3 py-2 border border-slate-200 hover:bg-slate-50"
            type="button"
          >
            إغلاق
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function TableHeaderCell({ children }) {
  return (
    <th className="px-4 py-3 text-right text-slate-600 font-bold whitespace-nowrap">
      {children}
    </th>
  );
}

function TableCell({ children }) {
  return <td className="px-4 py-4 align-middle">{children}</td>;
}

function sortByField(arr, field, dir) {
  const copy = [...arr];
  copy.sort((a, b) => {
    let av = a[field];
    let bv = b[field];
    if (typeof av === "string") av = av.trim();
    if (typeof bv === "string") bv = bv.trim();
    if (av < bv) return dir === "asc" ? -1 : 1;
    if (av > bv) return dir === "asc" ? 1 : -1;
    return 0;
  });
  return copy;
}

function getStageStatusSubtitle(project, stageIndex) {
  const cur = getCurrentStage(project);
  if (stageIndex !== cur) return null;
  const c = computeCounterDays(project);
  if (c < 0) return "(معلق)";
  return null;
}

function validateApprove(project, stageIndex) {
  // لا يمكن اعتماد مرحلة لاحقة قبل السابقة
  for (let i = 1; i < stageIndex; i++) {
    if (project.stages[i - 1].status !== "تم") {
      return `لا يمكن اعتماد ${stageLabel(stageIndex)} قبل اعتماد ${stageLabel(i)}.`;
    }
  }
  return null;
}

export default function MosqueProjectsTrackerApp() {
  const [projects, setProjects] = useState(seedProjects);

  useEffect(() => {
    // تحميل جميع المشاريع من ملف JSON (public/projects.json)
    // إذا لم يوجد الملف، سيبقى النظام على البيانات الافتراضية.
    loadProjectsFromJson(setProjects);
  }, []);
  const [q, setQ] = useState("");
  const [countryFilter, setCountryFilter] = useState("كل الدول");
  const [stageFilter, setStageFilter] = useState("كل المراحل");
  const [sortField, setSortField] = useState("project_no");
  const [sortDir, setSortDir] = useState("desc");
  const [modalOpen, setModalOpen] = useState(false);
  const [alert, setAlert] = useState(null);

  const stats = useMemo(() => computeStatusSummary(projects), [projects]);

  const countries = useMemo(() => {
    // قائمة ثابتة حسب طلبك
    return ["كل الدول", "السعودية", "عمان", "تركيا", "قطر"];
  }, []);

  const filtered = useMemo(() => {
    const query = q.trim();
    return projects.filter((p) => {
      const matchesQ =
        query.length === 0 ||
        String(p.project_no).includes(query) ||
        p.contractor_name.includes(query);

      const matchesCountry =
        countryFilter === "كل الدول" || p.country === countryFilter;

      const cur = getCurrentStage(p);
      const matchesStage =
        stageFilter === "كل المراحل" ||
        (stageFilter === "مكتمل" ? p.stages.every((s) => s.status === "تم") : String(cur) === stageFilter);

      return matchesQ && matchesCountry && matchesStage;
    });
  }, [projects, q, countryFilter, stageFilter]);

  const sorted = useMemo(() => {
    let base = filtered;

    if (sortField === "counter") {
      base = [...base].sort((a, b) => {
        const av = computeCounterDays(a);
        const bv = computeCounterDays(b);
        return sortDir === "asc" ? av - bv : bv - av;
      });
      return base;
    }

    return sortByField(base, sortField, sortDir);
  }, [filtered, sortField, sortDir]);

  function toggleSort(field) {
    if (sortField === field) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortField(field);
      setSortDir("asc");
    }
  }

  function approveStage(projectId, stageIndex) {
    setProjects((prev) => {
      const copy = prev.map((p) => ({
        ...p,
        stages: p.stages.map((s) => ({ ...s })),
      }));

      const p = copy.find((x) => x.id === projectId);
      if (!p) return prev;

      const msg = validateApprove(p, stageIndex);
      if (msg) {
        setAlert({ type: "error", message: msg });
        return prev;
      }

      // اعتماد
      p.stages[stageIndex - 1].status = "تم";
      p.stages[stageIndex - 1].approved_at = new Date();

      setAlert({
        type: "success",
        message: `تم اعتماد ${stageLabel(stageIndex)} للمشروع رقم ${p.project_no}.`,
      });

      return copy;
    });
  }

  function setStageRequired(projectId, stageIndex) {
    // رجوع للـ "مطلوب" (اختياري) مع مسح تاريخ الاعتماد
    setProjects((prev) => {
      const copy = prev.map((p) => ({
        ...p,
        stages: p.stages.map((s) => ({ ...s })),
      }));

      const p = copy.find((x) => x.id === projectId);
      if (!p) return prev;

      // لا تسمح بإرجاع مرحلة سابقة إن كانت مرحلة لاحقة معتمدة (لتجنب كسر التسلسل)
      for (let i = stageIndex + 1; i <= 4; i++) {
        if (p.stages[i - 1].status === "تم") {
          setAlert({
            type: "error",
            message: `لا يمكن إرجاع ${stageLabel(stageIndex)} إلى "مطلوب" لأن ${stageLabel(i)} معتمدة.`,
          });
          return prev;
        }
      }

      p.stages[stageIndex - 1].status = "مطلوب";
      p.stages[stageIndex - 1].approved_at = null;
      setAlert({
        type: "success",
        message: `تم تعديل ${stageLabel(stageIndex)} إلى "مطلوب" للمشروع رقم ${p.project_no}.`,
      });
      return copy;
    });
  }

  function addProject(newP) {
    setProjects((prev) => [newP, ...prev]);
    setAlert({ type: "success", message: `تمت إضافة مشروع رقم ${newP.project_no}.` });
  }

  return (
    <div className="min-h-screen bg-slate-50" dir="rtl">
      {/* Header */}
      <div className="bg-gradient-to-r from-blue-700 to-blue-500 text-white">
        <div className="max-w-6xl mx-auto px-5 py-10">
          <motion.h1
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-3xl md:text-4xl font-extrabold"
          >
            نظام متابعة مشاريع المساجد
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.05 }}
            className="mt-2 text-white/85"
          >
            إدارة المراحل، احتساب المدد الزمنية، ومتابعة التأخير التراكمي
          </motion.p>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-5 -mt-8 pb-16">
        {/* Stats */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
          <CardStat title="مشاريع منتهية (المرحلة 4)" value={stats.completed} accent="green" />
          <CardStat title="مشاريع متأخرة (بالسالب)" value={stats.late} accent="red" />
          <CardStat title="إجمالي المشاريع" value={stats.total} accent="blue" />
        </div>

        {/* Filters */}
        <div className="mt-6 bg-white rounded-2xl shadow-sm border border-slate-100 p-4">
          <div className="flex flex-col md:flex-row gap-3 md:items-center md:justify-between">
            <div className="flex gap-3 flex-col md:flex-row md:items-center">
              <Button onClick={() => setModalOpen(true)}>+ مشروع جديد</Button>
              <Select
                ariaLabel="فلترة الدول"
                value={countryFilter}
                onChange={setCountryFilter}
              >
                {countries.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </Select>

              <Select
                ariaLabel="فلترة المراحل"
                value={stageFilter}
                onChange={setStageFilter}
              >
                <option value="كل المراحل">كل المراحل</option>
                <option value="1">المرحلة 1</option>
                <option value="2">المرحلة 2</option>
                <option value="3">المرحلة 3</option>
                <option value="4">المرحلة 4</option>
                <option value="مكتمل">مكتمل</option>
              </Select>
            </div>

            <div className="flex gap-3 flex-col md:flex-row md:items-center">
              <Input value={q} onChange={setQ} placeholder="بحث باسم المقاول أو رقم المشروع..." />
              <Select
                ariaLabel="فرز"
                value={`${sortField}:${sortDir}`}
                onChange={(v) => {
                  const [f, d] = v.split(":");
                  setSortField(f);
                  setSortDir(d);
                }}
              >
                <option value="project_no:desc">الأحدث رقمًا</option>
                <option value="project_no:asc">الأقدم رقمًا</option>
                <option value="contractor_name:asc">اسم المقاول (أ-ي)</option>
                <option value="contractor_name:desc">اسم المقاول (ي-أ)</option>
                <option value="country:asc">الدولة (أ-ي)</option>
                <option value="counter:asc">الأكثر تأخرًا أولاً</option>
                <option value="counter:desc">الأكثر تقدّمًا أولاً</option>
              </Select>
            </div>
          </div>

          {alert ? (
            <div
              className={`mt-4 rounded-xl px-4 py-3 border text-sm font-semibold ${
                alert.type === "error"
                  ? "bg-rose-50 border-rose-200 text-rose-700"
                  : "bg-emerald-50 border-emerald-200 text-emerald-700"
              }`}
            >
              <div className="flex items-center justify-between gap-3">
                <div>{alert.message}</div>
                <button
                  type="button"
                  className="px-3 py-1 rounded-lg border border-current/20 hover:bg-white/40"
                  onClick={() => setAlert(null)}
                >
                  إخفاء
                </button>
              </div>
            </div>
          ) : null}
        </div>

        {/* Table */}
        <div className="mt-6 bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="min-w-full">
              <thead className="bg-slate-50 border-b border-slate-100">
                <tr>
                  <TableHeaderCell>
                    <button
                      className="hover:underline"
                      onClick={() => toggleSort("project_no")}
                      type="button"
                    >
                      رقم المشروع
                    </button>
                  </TableHeaderCell>
                  <TableHeaderCell>
                    <button
                      className="hover:underline"
                      onClick={() => toggleSort("contractor_name")}
                      type="button"
                    >
                      اسم المقاول
                    </button>
                  </TableHeaderCell>
                  <TableHeaderCell>المرحلة 1</TableHeaderCell>
                  <TableHeaderCell>المرحلة 2</TableHeaderCell>
                  <TableHeaderCell>المرحلة 3</TableHeaderCell>
                  <TableHeaderCell>المرحلة 4</TableHeaderCell>
                  <TableHeaderCell>
                    <button
                      className="hover:underline"
                      onClick={() => {
                        setSortField("counter");
                        setSortDir((d) => (d === "asc" ? "desc" : "asc"));
                      }}
                      type="button"
                    >
                      العداد (أيام)
                    </button>
                  </TableHeaderCell>
                  <TableHeaderCell>المدة</TableHeaderCell>
                  <TableHeaderCell>
                    <button
                      className="hover:underline"
                      onClick={() => toggleSort("country")}
                      type="button"
                    >
                      الدولة
                    </button>
                  </TableHeaderCell>
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-100">
                {sorted.map((p) => {
                  const counter = computeCounterDays(p);
                  const isLate = counter < 0;
                  const rowBg = isLate ? "bg-rose-50/60" : "bg-white";
                  const counterText = isLate
                    ? `(${Math.abs(counter)}-)`
                    : `${counter}+`;

                  return (
                    <tr key={p.id} className={`${rowBg} hover:bg-slate-50/60 transition`}>
                      <TableCell>
                        <div className="font-extrabold text-slate-900">{p.project_no}</div>
                      </TableCell>

                      <TableCell>
                        <div className="font-extrabold text-slate-900">{p.contractor_name}</div>
                      </TableCell>

                      {[1, 2, 3, 4].map((i) => {
                        const st = p.stages[i - 1].status;
                        const subtitle = getStageStatusSubtitle(p, i);
                        return (
                          <TableCell key={i}>
                            <div className="flex items-center gap-2">
                              {statusPill(st, subtitle)}
                              <div className="flex flex-col gap-2">
                                {st !== "تم" ? (
                                  <button
                                    type="button"
                                    className="text-xs font-bold text-blue-700 hover:underline"
                                    onClick={() => approveStage(p.id, i)}
                                  >
                                    اعتماد
                                  </button>
                                ) : (
                                  <button
                                    type="button"
                                    className="text-xs font-bold text-slate-600 hover:underline"
                                    onClick={() => setStageRequired(p.id, i)}
                                  >
                                    تعديل لمطلوب
                                  </button>
                                )}
                              </div>
                            </div>
                          </TableCell>
                        );
                      })}

                      <TableCell>
                        <div className={`font-extrabold ${isLate ? "text-rose-700" : "text-emerald-700"}`}>
                          {counterText}
                        </div>
                        <div className="text-xs text-slate-500 mt-1">
                          المرحلة الحالية: {stageLabel(getCurrentStage(p))}
                        </div>
                      </TableCell>

                      <TableCell>
                        <div className="text-slate-900 font-bold">
                          {formatDurationDaysToMonthsDays(p.duration_days)}
                        </div>
                        <div className="text-xs text-slate-500 mt-1">{p.duration_days} يوم</div>
                      </TableCell>

                      <TableCell>
                        <div className="text-slate-900 font-bold">{p.country}</div>
                      </TableCell>
                    </tr>
                  );
                })}

                {sorted.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="px-6 py-12 text-center text-slate-500">
                      لا توجد نتائج مطابقة.
                    </td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </div>

        {/* Footer note */}
        <div className="mt-6 text-sm text-slate-500">
          * العداد يُحسب للمرحلة الحالية فقط: إذا انتهت مدة المرحلة ولم تُعتمد يصبح بالسالب ويستمر حتى الاعتماد.
        </div>
      </div>

      <NewProjectModal
        open={modalOpen}
        onClose={() => setModalOpen(false)}
        onAdd={(p) => {
          addProject(p);
          setModalOpen(false);
        }}
      />
    </div>
  );
}

function NewProjectModal({ open, onClose, onAdd }) {
  const [projectNo, setProjectNo] = useState("");
  const [contractor, setContractor] = useState("");
  const [country, setCountry] = useState("السعودية");
  const [durationDays, setDurationDays] = useState("90");
  const [startDate, setStartDate] = useState(() => {
    const d = new Date();
    return d.toISOString().slice(0, 10);
  });
  const [err, setErr] = useState(null);

  function resetErr() {
    if (err) setErr(null);
  }

  function handleAdd() {
    const pno = Number(projectNo);
    const dd = Number(durationDays);

    if (!Number.isFinite(pno) || pno <= 0) {
      setErr("أدخل رقم مشروع صحيح.");
      return;
    }
    if (contractor.trim().length < 2) {
      setErr("أدخل اسم مقاول صحيح.");
      return;
    }
    if (!Number.isFinite(dd) || dd < 4) {
      setErr("أدخل مدة صحيحة (بالأيام). الحد الأدنى 4 أيام.");
      return;
    }

    const newP = {
      id: `p-${pno}-${Math.random().toString(16).slice(2)}`,
      project_no: pno,
      contractor_name: contractor.trim(),
      country,
      duration_days: clampInt(dd, 4, 3650),
      start_date: new Date(startDate),
      stages: [
        { status: "مطلوب", approved_at: null },
        { status: "مطلوب", approved_at: null },
        { status: "مطلوب", approved_at: null },
        { status: "مطلوب", approved_at: null },
      ],
    };

    onAdd(newP);
    setProjectNo("");
    setContractor("");
    setCountry("السعودية");
    setDurationDays("90");
    const d = new Date();
    setStartDate(d.toISOString().slice(0, 10));
    setErr(null);
  }

  return (
    <Modal open={open} title="إضافة مشروع جديد" onClose={onClose}>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <div className="text-sm font-bold text-slate-700 mb-2">رقم المشروع</div>
          <input
            className="w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            value={projectNo}
            onChange={(e) => {
              setProjectNo(e.target.value);
              resetErr();
            }}
            placeholder="مثال: 1004"
          />
        </div>

        <div>
          <div className="text-sm font-bold text-slate-700 mb-2">اسم المقاول</div>
          <input
            className="w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            value={contractor}
            onChange={(e) => {
              setContractor(e.target.value);
              resetErr();
            }}
            placeholder="مثال: مؤسسة البناء"
          />
        </div>

        <div>
          <div className="text-sm font-bold text-slate-700 mb-2">الدولة</div>
          <select
            className="w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            value={country}
            onChange={(e) => {
              setCountry(e.target.value);
              resetErr();
            }}
          >
            <option value="السعودية">السعودية</option>
            <option value="عمان">عمان</option>
            <option value="السعودية">السعودية</option>
            <option value="عمان">عمان</option>
            <option value="تركيا">تركيا</option>
            <option value="قطر">قطر</option>
          </select>
        </div>

        <div>
          <div className="text-sm font-bold text-slate-700 mb-2">المدة (بالأيام)</div>
          <input
            className="w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            value={durationDays}
            onChange={(e) => {
              setDurationDays(e.target.value);
              resetErr();
            }}
            placeholder="مثال: 120"
          />
          <div className="text-xs text-slate-500 mt-2">
            سيتم تقسيمها تلقائيًا إلى 4 مراحل (25% لكل مرحلة).
          </div>
        </div>

        <div className="md:col-span-2">
          <div className="text-sm font-bold text-slate-700 mb-2">تاريخ بداية المشروع</div>
          <input
            type="date"
            className="w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-slate-800 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-200"
            value={startDate}
            onChange={(e) => {
              setStartDate(e.target.value);
              resetErr();
            }}
          />
        </div>
      </div>

      {err ? (
        <div className="mt-4 rounded-xl px-4 py-3 border bg-rose-50 border-rose-200 text-rose-700 text-sm font-semibold">
          {err}
        </div>
      ) : null}

      <div className="mt-5 flex gap-3">
        <Button onClick={handleAdd}>حفظ</Button>
        <Button variant="ghost" onClick={onClose}>
          إلغاء
        </Button>
      </div>
    </Modal>
  );
}
