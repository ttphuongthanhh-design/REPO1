import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  Plus, 
  Trash2, 
  Edit2, 
  Calendar, 
  Search, 
  Check, 
  X, 
  AlertCircle, 
  Clock, 
  TrendingUp, 
  Users, 
  Percent, 
  FileText, 
  Download, 
  ChevronLeft, 
  ChevronRight,
  Sparkles,
  Sun,
  Moon,
  Link as LinkIcon,
  Lock,
  Unlock,
  Settings
} from 'lucide-react';

// Interfaces
interface Task {
  id: number;
  title: string;
  desc: string;
  scope: 'ae' | 'si' | 'pd' | 'va' | 'pr';
  subtype: string;
  assignee: string;
  priority: 'high' | 'med' | 'low';
  col: number; // 0: Backlog, 1: In Progress, 2: Review, 3: Done, 4: Failed/Reject
  start: string;
  deadline: string;
  pct: number;
  note: string;
  link?: string;
  completedAt?: string; // ISO date set when the task is marked Done (col 3)
  updatedAt?: string;   // ISO timestamp of the last change (drives the Task Tracker "last update")
}

// One entry in the Activity Log — the audit trail that lets you follow a task end-to-end.
interface Activity {
  at: string;       // ISO timestamp
  taskId: number;
  title: string;    // task title at the time of the event
  type: 'created' | 'status' | 'progress' | 'done' | 'edited' | 'deleted';
  detail: string;   // human-readable summary, e.g. "Backlog → In Progress"
}

interface Member {
  id: string;
  name: string;
  role: string;
  color: string;
  removable?: boolean;
}

interface ScopeConfig {
  name: string;
  color: string;
  glow: string;
  badge: string;
  text: string;
}

// Constant Constants
const COLS = ['Backlog', 'In Progress', 'Review', 'Done', 'Failed / Reject'];
// Only the Host Lead (id 'T') is a locked default; everyone else is editable/removable.
const LOCKED_MEMBER_ID = 'T';
const DEFAULT_MEMBERS: Member[] = [
  { id: 'T', name: 'TIFFANY', role: 'Senior — Host Lead', color: 'linear-gradient(135deg, #818cf8, #4f46e5)', removable: false },
  { id: 'M', name: 'MERCURY', role: 'Production Executive', color: 'linear-gradient(135deg, #f43f5e, #be123c)', removable: true },
];

// Preset gradient palette for new member avatars
const MEMBER_COLORS = [
  'linear-gradient(135deg, #818cf8, #4f46e5)',
  'linear-gradient(135deg, #f43f5e, #be123c)',
  'linear-gradient(135deg, #34d399, #059669)',
  'linear-gradient(135deg, #fbbf24, #d97706)',
  'linear-gradient(135deg, #22d3ee, #0891b2)',
  'linear-gradient(135deg, #f472b6, #db2777)',
  'linear-gradient(135deg, #c084fc, #7c3aed)',
  'linear-gradient(135deg, #94a3b8, #475569)',
];
const SCOPES: Record<'ae' | 'si' | 'pd' | 'va' | 'pr', string> = {
  ae: 'Activation & Event',
  si: 'Social & Influencers',
  pd: 'Performance & Data',
  va: 'Visual & AI',
  pr: 'Production & Operation'
};

const SCOPE_CONFIGS: Record<'ae' | 'si' | 'pd' | 'va' | 'pr', ScopeConfig> = {
  ae: { 
    name: 'Activation & Event', 
    color: '#8b5cf6', 
    glow: 'rgba(139, 92, 246, 0.25)', 
    badge: 'bg-violet-950/40 text-violet-400 border border-violet-800/30',
    text: 'text-violet-400'
  },
  si: { 
    name: 'Social & Influencers', 
    color: '#ec4899', 
    glow: 'rgba(236, 72, 153, 0.25)', 
    badge: 'bg-pink-950/40 text-pink-400 border border-pink-800/30',
    text: 'text-pink-400'
  },
  pd: {
    name: 'Performance & Data',
    color: '#06b6d4',
    glow: 'rgba(6, 182, 212, 0.25)',
    badge: 'bg-cyan-950/40 text-cyan-400 border border-cyan-800/30',
    text: 'text-cyan-400'
  },
  va: {
    name: 'Visual & AI',
    color: '#14b8a6',
    glow: 'rgba(20, 184, 166, 0.25)',
    badge: 'bg-teal-950/40 text-teal-400 border border-teal-800/30',
    text: 'text-teal-400'
  },
  pr: {
    name: 'Production & Operation',
    color: '#f43f5e',
    glow: 'rgba(244, 63, 94, 0.25)', 
    badge: 'bg-rose-950/40 text-rose-400 border border-rose-800/30',
    text: 'text-rose-400'
  }
};

const COL_COLORS = ['#94a3b8', '#f59e0b', '#3b82f6', '#10b981', '#ef4444'];
const PRIORITY_BADGES = {
  high: 'bg-red-950/40 text-red-400 border border-red-800/30',
  med: 'bg-amber-950/40 text-amber-400 border border-amber-800/30',
  low: 'bg-blue-950/40 text-blue-400 border border-blue-800/30'
};

const defaultSubtypes: Record<'ae' | 'si' | 'pd' | 'va' | 'pr', string[]> = {
  ae: ['Planning & Brief', 'Concept Development', 'Vendor Management', 'Site Survey', 'Timeline & Budget', 'Proposal / Sale Kit', 'Post-event Report'],
  si: ['Influencer Scouting', 'Influencer Brief', 'Content Plan', 'KOL Coordination', 'Performance Tracking', 'Community Management'],
  pd: ['Data Collection', 'Dashboard & Report', 'Audience Segmentation', 'Performance Tracking', 'Insight Analysis', 'Reporting & KPI'],
  va: ['Visual Concept', 'Design / Artwork', 'AI Tool / Automation', 'Video & Motion', 'Template & System', 'Asset Library'],
  pr: ['Site Survey', 'Quotation & Supplier', 'Production Monitoring', 'POSM / Print', 'Logistics & Setup', 'PG/PB Management', 'Onsite Supervision', 'Post-production Report', 'Checklist & QC']
};

const GANTT_DAYS = 42; // 6 weeks view

// ===== Target KPIs — Annual Revenue Planning by Distribution Channel =====
interface KpiChannel {
  id: number;
  channel: string;
  allocation: number;            // % of total annual target (user-editable)
  priority: 'High' | 'Medium' | 'Low'; // user-editable
  notes: string;                 // user-editable
  characteristics: string;       // channel description (Section 2)
  curve: [number, number, number, number]; // fixed seasonal split across Q1..Q4 (sums to 1)
  actual?: number[];             // 12 monthly actual-KPI figures (user-editable)
}
interface KpiYear {
  totalTarget: number;           // total annual target in VND (user-editable, prominent)
  channels: KpiChannel[];
}
interface KpiData {
  years: Record<string, KpiYear>; // data per planning year (2026/2027/2028…)
}
const KPI_YEARS = ['2026', '2027', '2028'];
// 7 monochrome indigo tones (dark → light) to distinguish channels on stacked columns
const CH_TONES = ['#312e81', '#3730a3', '#4338ca', '#4f46e5', '#6366f1', '#818cf8', '#a5b4fc'];
const BASE_KPI_CHANNELS: KpiChannel[] = [
  { id: 1, channel: 'University',      allocation: 20, priority: 'High',   notes: 'Back-to-school & exam seasons drive H2 demand', characteristics: 'Seasonal, term-driven; strong H2 ramp', curve: [0.20, 0.22, 0.28, 0.30] },
  { id: 2, channel: 'Fast Food',       allocation: 18, priority: 'High',   notes: 'Stable year-round footfall',                   characteristics: 'High frequency, low seasonality',     curve: [0.25, 0.25, 0.25, 0.25] },
  { id: 3, channel: 'Coffee Shop',     allocation: 16, priority: 'Medium', notes: 'Front-loaded; cools toward year-end',          characteristics: 'Front-loaded; H1 heavy',              curve: [0.28, 0.26, 0.24, 0.22] },
  { id: 4, channel: 'Salon',           allocation: 12, priority: 'Medium', notes: 'Gradual ramp toward festive season',           characteristics: 'Steady ramp; festive uplift',         curve: [0.22, 0.24, 0.26, 0.28] },
  { id: 5, channel: 'Building',        allocation: 14, priority: 'Low',    notes: 'Project-based; concentrated early in year',    characteristics: 'Project-based; H1 concentrated',      curve: [0.30, 0.28, 0.22, 0.20] },
  { id: 6, channel: 'Apartment',       allocation: 12, priority: 'Medium', notes: 'Even spread, slight H2 lift',                  characteristics: 'Residential; balanced demand',        curve: [0.24, 0.24, 0.26, 0.26] },
  { id: 7, channel: 'Hotel & Resort',  allocation: 8,  priority: 'High',   notes: 'Peaks mid-year (summer travel)',               characteristics: 'Tourism-led; mid-year peak',          curve: [0.20, 0.30, 0.30, 0.20] },
];
const mkKpiYear = (target: number): KpiYear => ({ totalTarget: target, channels: BASE_KPI_CHANNELS.map(c => ({ ...c })) });
const DEFAULT_KPIS: KpiData = {
  years: { '2026': mkKpiYear(50_000_000_000), '2027': mkKpiYear(57_500_000_000), '2028': mkKpiYear(66_000_000_000) },
};
// Accept both the new per-year shape and the legacy { totalTarget, channels } shape.
const migrateKpis = (raw: any): KpiData => {
  if (raw && raw.years && typeof raw.years === 'object') return raw;
  if (raw && Array.isArray(raw.channels)) {
    const base = raw.channels;
    const t = raw.totalTarget || 50_000_000_000;
    return { years: { '2026': { totalTarget: t, channels: base }, '2027': mkKpiYear(Math.round(t * 1.15)), '2028': mkKpiYear(Math.round(t * 1.32)) } };
  }
  return DEFAULT_KPIS;
};

// --- Security: simple obfuscating hash + baked default edit password ---
const hashPwd = (s: string) => {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
  return 'h' + h.toString(36);
};
// Default edit password baked into the file (travels with the file when shared).
const BAKED_EDIT_HASH = hashPwd('G@M1234');

const defaultTasks: Task[] = [
  { id: 1, title: 'Brief intake & concept — Q3 Activation', desc: '', scope: 'ae', subtype: 'Concept Development', assignee: 'T', priority: 'high', col: 1, start: '2025-07-01', deadline: '2025-07-08', pct: 40, note: 'Đang hoàn thiện mood board' },
  { id: 2, title: 'Vendor quotation — summer pop-up', desc: '', scope: 'pr', subtype: 'Quotation & Supplier', assignee: 'M', priority: 'high', col: 1, start: '2025-07-02', deadline: '2025-07-10', pct: 30, note: 'Đã liên hệ 3/6 vendor' },
  { id: 3, title: 'Site survey — Quận 1', desc: '', scope: 'pr', subtype: 'Site Survey', assignee: 'M', priority: 'high', col: 2, start: '2025-07-01', deadline: '2025-07-05', pct: 80, note: 'Đã survey, đang viết báo cáo' },
  { id: 4, title: 'Timeline & logistics plan Q3', desc: '', scope: 'ae', subtype: 'Timeline & Budget', assignee: 'T', priority: 'high', col: 1, start: '2025-07-03', deadline: '2025-07-12', pct: 50, note: 'Đang draft timeline' },
  { id: 5, title: 'KOL list — July campaign', desc: '', scope: 'si', subtype: 'Influencer Scouting', assignee: 'T', priority: 'med', col: 0, start: '2025-07-07', deadline: '2025-07-15', pct: 10, note: '' },
  { id: 6, title: 'Content plan — Social July', desc: '', scope: 'si', subtype: 'Content Plan', assignee: 'T', priority: 'high', col: 2, start: '2025-06-25', deadline: '2025-07-01', pct: 75, note: 'Chờ client approve' },
  { id: 7, title: 'Performance dashboard — KPI', desc: '', scope: 'pd', subtype: 'Dashboard & Report', assignee: 'T', priority: 'high', col: 1, start: '2025-07-01', deadline: '2025-07-03', pct: 60, note: 'Đang pull data từ Meta' },
  { id: 8, title: 'Audience segmentation Q2', desc: '', scope: 'pd', subtype: 'Audience Segmentation', assignee: 'T', priority: 'low', col: 3, start: '2025-06-20', deadline: '2025-06-30', pct: 100, note: 'Hoàn thành' },
  { id: 9, title: 'POSM production monitoring', desc: '', scope: 'pr', subtype: 'Production Monitoring', assignee: 'M', priority: 'high', col: 1, start: '2025-07-03', deadline: '2025-07-08', pct: 45, note: 'Đang theo dõi tại xưởng' },
  { id: 10, title: 'PG/PB recruitment & briefing', desc: '', scope: 'pr', subtype: 'PG/PB Management', assignee: 'M', priority: 'med', col: 0, start: '2025-07-10', deadline: '2025-07-18', pct: 0, note: '' },
  { id: 11, title: 'AI reporting template', desc: '', scope: 'va', subtype: 'AI Tool / Automation', assignee: 'T', priority: 'low', col: 3, start: '2025-06-22', deadline: '2025-06-28', pct: 100, note: 'Done' },
  { id: 12, title: 'Sale kit & proposal — Activation', desc: '', scope: 'ae', subtype: 'Proposal / Sale Kit', assignee: 'T', priority: 'med', col: 0, start: '2025-07-14', deadline: '2025-07-22', pct: 0, note: '' },
  { id: 13, title: 'Logistics & setup plan — onsite', desc: '', scope: 'pr', subtype: 'Logistics & Setup', assignee: 'M', priority: 'med', col: 1, start: '2025-07-05', deadline: '2025-07-14', pct: 20, note: 'Đang liên hệ vendor vận chuyển' },
  { id: 14, title: 'Checklist QC — pre-event', desc: '', scope: 'pr', subtype: 'Checklist & QC', assignee: 'M', priority: 'high', col: 0, start: '2025-07-12', deadline: '2025-07-19', pct: 0, note: '' },
  { id: 15, title: 'Influencer brief & contract', desc: '', scope: 'si', subtype: 'Influencer Brief', assignee: 'T', priority: 'med', col: 0, start: '2025-07-12', deadline: '2025-07-20', pct: 0, note: '' }
];

export default function App() {
  // --- States ---
  const [tasks, setTasks] = useState<Task[]>(() => {
    try {
      const saved = localStorage.getItem('impact_tasks');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Migrate removed scope 'da' (Data & AI) -> 'pd' (Performance & Data)
          return parsed.map((t: Task) => (t.scope as string) === 'da' ? { ...t, scope: 'pd' } : t);
        }
      }
    } catch (e) {
      console.error(e);
    }
    return defaultTasks;
  });

  const [subtypes, setSubtypes] = useState<Record<'ae' | 'si' | 'pd' | 'va' | 'pr', string[]>>(() => {
    try {
      const saved = localStorage.getItem('impact_subtypes');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          // Migrate 'da' subtypes -> 'pd'; ensure 'pd' and 'va' keys exist
          if (parsed.da && !parsed.pd) { parsed.pd = parsed.da; }
          delete parsed.da;
          if (!parsed.pd) parsed.pd = defaultSubtypes.pd;
          if (!parsed.va) parsed.va = defaultSubtypes.va;
          return parsed;
        }
      }
    } catch (e) {
      console.error(e);
    }
    return defaultSubtypes;
  });

  const [members, setMembers] = useState<Member[]>(() => {
    try {
      const saved = localStorage.getItem('impact_members');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Migration: only the Host Lead is locked; everyone else is editable/removable.
          return parsed.map((m: Member) => ({ ...m, removable: m.id !== LOCKED_MEMBER_ID }));
        }
      }
    } catch (e) {
      console.error(e);
    }
    return DEFAULT_MEMBERS;
  });

  // --- Activity Log (audit trail powering the Task Tracker) ---
  const [activity, setActivity] = useState<Activity[]>(() => {
    try {
      const saved = localStorage.getItem('impact_activity');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (e) {
      console.error(e);
    }
    return [];
  });

  // Activity Log retention: keep entries for 3 months; older ones auto-expire.
  const ACTIVITY_RETENTION_MONTHS = 3;
  const isWithinRetention = (iso: string) => {
    if (!iso) return false;
    const cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - ACTIVITY_RETENTION_MONTHS);
    return new Date(iso).getTime() >= cutoff.getTime();
  };

  const ACTIVITY_LIMIT = 30; // log holds max 30; the 31st resets the counter back to 1 and loops
  const logActivity = (taskId: number, title: string, type: Activity['type'], detail: string) => {
    const entry: Activity = { at: new Date().toISOString(), taskId, title, type, detail };
    setActivity(prev => {
      const next = [entry, ...prev].filter(a => isWithinRetention(a.at));
      return next.length > ACTIVITY_LIMIT ? [entry] : next; // hit 31 → reset to 1
    });
  };

  useEffect(() => {
    try { localStorage.setItem('impact_activity', JSON.stringify(activity)); } catch (e) { console.error(e); }
  }, [activity]);

  // Prune entries older than the retention window on load (covers data restored from storage/server).
  useEffect(() => {
    setActivity(prev => {
      const kept = prev.filter(a => isWithinRetention(a.at)).slice(0, ACTIVITY_LIMIT);
      return kept.length === prev.length ? prev : kept;
    });
  }, []);

  // --- Target KPIs (annual revenue planning, per year) ---
  const [kpis, setKpis] = useState<KpiData>(() => {
    try {
      const saved = localStorage.getItem('impact_kpis');
      if (saved) {
        const parsed = JSON.parse(saved);
        const m = migrateKpis(parsed);
        if (m && m.years && Object.keys(m.years).length > 0) return m;
      }
    } catch (e) {
      console.error(e);
    }
    return DEFAULT_KPIS;
  });
  useEffect(() => {
    try { localStorage.setItem('impact_kpis', JSON.stringify(kpis)); } catch (e) { console.error(e); }
  }, [kpis]);
  const [kpiYear, setKpiYear] = useState<string>(KPI_YEARS[0]);
  // Channel filter (which distribution channels are shown), shared by both sections
  const [kpiChannelFilter, setKpiChannelFilter] = useState<number[] | null>(null); // null = all
  const [kpiChanMenuOpen, setKpiChanMenuOpen] = useState(false);
  const [kpiHoverYear, setKpiHoverYear] = useState<number | null>(null); // combo-chart tooltip
  const [comboView, setComboView] = useState<'year' | 'month'>('year');
  const [dashYear, setDashYear] = useState<string>(KPI_YEARS[0]);
  const kpiYearData: KpiYear = kpis.years[kpiYear] || kpis.years[KPI_YEARS[0]] || Object.values(kpis.years)[0];

  const updateKpiChannel = (id: number, field: keyof KpiChannel, value: any) => {
    if (!isEditMode) return;
    setKpis(prev => {
      const yd = prev.years[kpiYear];
      if (!yd) return prev;
      return { ...prev, years: { ...prev.years, [kpiYear]: { ...yd, channels: yd.channels.map(c => c.id === id ? { ...c, [field]: field === 'allocation' ? Math.max(0, +value || 0) : value } : c) } } };
    });
  };
  const updateKpiTotal = (value: any) => {
    if (!isEditMode) return;
    setKpis(prev => {
      const yd = prev.years[kpiYear];
      if (!yd) return prev;
      return { ...prev, years: { ...prev.years, [kpiYear]: { ...yd, totalTarget: Math.max(0, +value || 0) } } };
    });
  };
  const updateKpiActual = (id: number, monthIdx: number, value: any) => {
    if (!isEditMode) return;
    setKpis(prev => {
      const yd = prev.years[kpiYear];
      if (!yd) return prev;
      return { ...prev, years: { ...prev.years, [kpiYear]: { ...yd, channels: yd.channels.map(c => {
        if (c.id !== id) return c;
        const actual = Array.from({ length: 12 }, (_, i) => (c.actual && c.actual[i]) || 0);
        actual[monthIdx] = Math.max(0, +value || 0);
        return { ...c, actual };
      }) } } };
    });
  };
  // Monthly Target Allocation table — sort
  const [kpiMonthSort, setKpiMonthSort] = useState<{ col: 'name' | 'total' | number; dir: 'asc' | 'desc' }>({ col: 'total', dir: 'desc' });
  const toggleKpiMonthSort = (col: 'name' | 'total' | number) => {
    setKpiMonthSort(prev => prev.col === col ? { col, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { col, dir: typeof col === 'string' && col === 'name' ? 'asc' : 'desc' });
  };

  // Task Tracker table sorting
  const [trackerSort, setTrackerSort] = useState<{ field: 'title' | 'assignee' | 'scope' | 'col' | 'pct' | 'deadline' | 'updatedAt'; dir: 'asc' | 'desc' }>({ field: 'updatedAt', dir: 'desc' });
  const toggleTrackerSort = (field: typeof trackerSort.field) => {
    setTrackerSort(prev => prev.field === field ? { field, dir: prev.dir === 'asc' ? 'desc' : 'asc' } : { field, dir: 'asc' });
  };

  // --- Server sync (Postgres-backed via /api/state) ---
  // The app keeps working offline via localStorage; this layer adds shared persistence.
  const serverLoadedRef = React.useRef(false);

  useEffect(() => {
    let cancelled = false;
    fetch('/api/state')
      .then(r => (r.ok ? r.json() : null))
      .then(data => {
        if (cancelled || !data) return;
        const hasTasks = Array.isArray(data.tasks) && data.tasks.length > 0;
        if (hasTasks) {
          setTasks(data.tasks.map((t: Task) => (t.scope as string) === 'da' ? { ...t, scope: 'pd' } : t));
        }
        if (data.subtypes && typeof data.subtypes === 'object') {
          setSubtypes(prev => ({ ...prev, ...data.subtypes }));
        }
        if (Array.isArray(data.members) && data.members.length > 0) {
          setMembers(data.members.map((m: Member) => ({ ...m, removable: m.id !== LOCKED_MEMBER_ID })));
        }
        if (Array.isArray(data.activity)) {
          setActivity(data.activity.filter((a: Activity) => isWithinRetention(a.at)).slice(0, ACTIVITY_LIMIT));
        }
        if (data.kpis && (data.kpis.years || Array.isArray(data.kpis.channels))) {
          setKpis(migrateKpis(data.kpis));
        }
      })
      .catch(() => { /* offline: fall back to localStorage data already loaded */ })
      .finally(() => { if (!cancelled) serverLoadedRef.current = true; });
    return () => { cancelled = true; };
  }, []);

  // Persist combined state to the server (debounced) after the initial load.
  useEffect(() => {
    if (!serverLoadedRef.current) return;
    const id = setTimeout(() => {
      fetch('/api/state', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tasks, subtypes, members, activity, kpis }),
      }).catch(() => { /* ignore: localStorage still holds the data */ });
    }, 600);
    return () => clearTimeout(id);
  }, [tasks, subtypes, members, activity, kpis]);

  // Member lookup helpers
  const memberById = (id: string) => members.find(m => m.id === id);
  const memberName = (id: string) => memberById(id)?.name ?? id;
  const memberRole = (id: string) => memberById(id)?.role ?? '';
  const memberColor = (id: string) => memberById(id)?.color ?? 'linear-gradient(135deg, #94a3b8, #475569)';
  const memberInitial = (id: string) => (memberById(id)?.name ?? id).trim().charAt(0).toUpperCase() || '?';

  const [activeTab, setActiveTab] = useState<'dash' | 'kanban' | 'tracker' | 'timeline' | 'daily' | 'weekly' | 'kpis'>('dash');

  const [darkMode, setDarkMode] = useState(() => {
    try {
      const saved = localStorage.getItem('impact_darkMode');
      return saved === 'true';
    } catch (e) {
      return false;
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('impact_darkMode', String(darkMode));
      if (darkMode) {
        document.documentElement.classList.add('dark');
      } else {
        document.documentElement.classList.remove('dark');
      }
    } catch (e) {
      console.error(e);
    }
  }, [darkMode]);

  // Filter States
  const [searchQuery, setSearchQuery] = useState('');
  const [assigneeFilter, setAssigneeFilter] = useState<string>('all');
  const [scopeFilter, setScopeFilter] = useState<'all' | 'ae' | 'si' | 'pd' | 'va' | 'pr'>('all');
  const [timelineScopeFilter, setTimelineScopeFilter] = useState<'all' | 'ae' | 'si' | 'pd' | 'va' | 'pr'>('all');

  // Date States
  const [ganttStart, setGanttStart] = useState<Date>(() => {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - 14);
    return d;
  });
  const [dailyDate, setDailyDate] = useState(() => new Date().toISOString().split('T')[0]);
  const [dailyFilterMode, setDailyFilterMode] = useState<'date' | 'all'>('date');
  // Weekly Report filter: a Monday key (yyyy-mm-dd) of the selected week, or 'all'
  const [weeklyFilter, setWeeklyFilter] = useState<string>(() => {
    const d = new Date();
    const startDay = (Math.ceil(d.getDate() / 7) - 1) * 7 + 1;
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(startDay)}`;
  });
  const [weeklyNote, setWeeklyNote] = useState('');

  // --- Security / permissions ---
  // Fixed edit password baked into the file (G@M1234).
  const editPwdHash = BAKED_EDIT_HASH;
  // Default = View (locked). Owner unlocks with the password.
  const [isEditMode, setIsEditMode] = useState<boolean>(false);
  // Unlock (view -> edit) modal
  const [showUnlockModal, setShowUnlockModal] = useState(false);
  const [unlockInput, setUnlockInput] = useState('');
  const [unlockError, setUnlockError] = useState(false);

  const enterEditMode = () => {
    setUnlockInput('');
    setUnlockError(false);
    setShowUnlockModal(true);
  };

  const submitUnlock = () => {
    if (hashPwd(unlockInput) === editPwdHash) {
      setIsEditMode(true);
      setShowUnlockModal(false);
      setUnlockInput('');
      setUnlockError(false);
    } else {
      setUnlockError(true);
    }
  };

  const lockToView = () => setIsEditMode(false);

  // Modal States
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editId, setEditId] = useState<number | null>(null);
  const [taskForm, setTaskForm] = useState<Omit<Task, 'id'>>({
    title: '',
    desc: '',
    scope: 'ae',
    subtype: '',
    assignee: 'T',
    priority: 'high',
    col: 0,
    start: '',
    deadline: '',
    pct: 0,
    note: '',
    link: ''
  });

  const [draggedTaskId, setDraggedTaskId] = useState<number | null>(null);
  const [expandedTaskId, setExpandedTaskId] = useState<number | null>(null); // card whose details are open
  const [newSubtypeInput, setNewSubtypeInput] = useState('');
  const [isAddingSubtype, setIsAddingSubtype] = useState(false);

  // Member management modal
  const [isMemberModalOpen, setIsMemberModalOpen] = useState(false);
  const [editingMemberId, setEditingMemberId] = useState<string | null>(null);
  const [memberForm, setMemberForm] = useState<{ name: string; role: string; color: string }>({
    name: '',
    role: '',
    color: MEMBER_COLORS[2],
  });

  // --- Effects ---
  useEffect(() => {
    localStorage.setItem('impact_tasks', JSON.stringify(tasks));
  }, [tasks]);

  useEffect(() => {
    localStorage.setItem('impact_subtypes', JSON.stringify(subtypes));
  }, [subtypes]);

  useEffect(() => {
    localStorage.setItem('impact_members', JSON.stringify(members));
  }, [members]);

  // --- Member Management ---
  const resetMemberForm = () => {
    const used = new Set(members.map(m => m.color));
    const nextColor = MEMBER_COLORS.find(c => !used.has(c)) || MEMBER_COLORS[members.length % MEMBER_COLORS.length];
    setEditingMemberId(null);
    setMemberForm({ name: '', role: '', color: nextColor });
  };

  const openMemberModal = () => {
    if (!isEditMode) return;
    resetMemberForm();
    setIsMemberModalOpen(true);
  };

  const startEditMember = (mem: Member) => {
    setEditingMemberId(mem.id);
    setMemberForm({ name: mem.name, role: mem.role, color: mem.color });
  };

  const handleAddMember = () => {
    const name = memberForm.name.trim();
    if (!name) return alert('Vui lòng nhập tên thành viên.');
    // Build a unique id from the first letter, falling back to suffixes
    const base = name.charAt(0).toUpperCase();
    let id = base;
    let n = 2;
    while (members.some(m => m.id === id)) {
      id = base + n;
      n++;
    }
    setMembers(prev => [...prev, { id, name, role: memberForm.role.trim(), color: memberForm.color, removable: true }]);
    resetMemberForm();
  };

  const handleUpdateMember = () => {
    if (!editingMemberId) return;
    const name = memberForm.name.trim();
    if (!name) return alert('Vui lòng nhập tên thành viên.');
    setMembers(prev => prev.map(m =>
      m.id === editingMemberId
        ? { ...m, name, role: memberForm.role.trim(), color: memberForm.color }
        : m
    ));
    resetMemberForm();
  };

  const handleSubmitMember = () => {
    if (editingMemberId) handleUpdateMember();
    else handleAddMember();
  };

  const handleDeleteMember = (id: string) => {
    const target = memberById(id);
    if (!target || target.removable === false) return;
    const remaining = members.filter(m => m.id !== id);
    const fallback = remaining[0]?.id ?? '';
    const assignedCount = tasks.filter(t => t.assignee === id).length;
    const msg = assignedCount > 0
      ? `Xóa "${target.name}"? ${assignedCount} task sẽ được chuyển sang "${memberName(fallback)}".`
      : `Xóa "${target.name}"?`;
    if (!confirm(msg)) return;
    if (assignedCount > 0 && fallback) {
      setTasks(prev => prev.map(t => t.assignee === id ? { ...t, assignee: fallback } : t));
    }
    if (assigneeFilter === id) setAssigneeFilter('all');
    if (editingMemberId === id) resetMemberForm();
    setMembers(remaining);
  };

  // --- Helpers ---
  const handleOpenAddModal = (colIndex: number) => {
    if (!isEditMode) return;
    setEditId(null);
    setTaskForm({
      title: '',
      desc: '',
      scope: 'ae',
      subtype: defaultSubtypes.ae[0] || '',
      assignee: 'T',
      priority: 'high',
      col: colIndex,
      start: '',
      deadline: '',
      pct: 0,
      note: '',
      link: ''
    });
    setIsAddingSubtype(false);
    setNewSubtypeInput('');
    setIsModalOpen(true);
  };

  const handleOpenEditModal = (task: Task) => {
    if (!isEditMode) return;
    setEditId(task.id);
    setTaskForm({ ...task });
    setIsAddingSubtype(false);
    setNewSubtypeInput('');
    setIsModalOpen(true);
  };

  const closeModal = () => {
    setIsModalOpen(false);
    setEditId(null);
  };

  const handleDeleteTask = (id: number) => {
    if (!isEditMode) return;
    if (confirm('Delete this task?')) {
      const victim = tasks.find(t => t.id === id);
      setTasks(prev => prev.filter(t => t.id !== id));
      if (victim) logActivity(id, victim.title, 'deleted', `Đã xoá khỏi ${COLS[victim.col]}`);
      if (isModalOpen && editId === id) setIsModalOpen(false);
    }
  };

  const handleSaveTask = (e: React.FormEvent) => {
    e.preventDefault();
    if (!taskForm.title.trim()) return alert('Please enter a task title.');

    const formattedForm = { ...taskForm };
    if (formattedForm.pct === 100) formattedForm.col = 3;
    if (formattedForm.col === 3) formattedForm.pct = 100;
    // Stamp/clear the completion date so Team Velocity reflects real activity
    if (formattedForm.col === 3) {
      if (!formattedForm.completedAt) formattedForm.completedAt = new Date().toISOString().split('T')[0];
    } else {
      formattedForm.completedAt = undefined;
    }
    formattedForm.updatedAt = new Date().toISOString();

    if (editId !== null) {
      const before = tasks.find(t => t.id === editId);
      setTasks(prev => prev.map(t => t.id === editId ? { ...t, ...formattedForm } : t));
      if (before) {
        if (before.col !== formattedForm.col) {
          logActivity(editId, formattedForm.title, formattedForm.col === 3 ? 'done' : 'status', `${COLS[before.col]} → ${COLS[formattedForm.col]}`);
        } else {
          logActivity(editId, formattedForm.title, 'edited', 'Cập nhật thông tin task');
        }
      }
    } else {
      const nextId = tasks.reduce((max, t) => Math.max(max, t.id), 0) + 1;
      setTasks(prev => [...prev, { id: nextId, ...formattedForm }]);
      logActivity(nextId, formattedForm.title, 'created', `Tạo mới trong ${COLS[formattedForm.col]}`);
    }
    setIsModalOpen(false);
  };

  const updateTaskField = (id: number, field: keyof Task, value: any) => {
    if (!isEditMode) return;
    const before = tasks.find(t => t.id === id);
    setTasks(prev => prev.map(t => {
      if (t.id === id) {
        const updated = { ...t, [field]: value };
        if (field === 'pct') {
          updated.pct = Math.min(100, Math.max(0, +value || 0));
          if (updated.pct === 100) updated.col = 3;
        } else if (field === 'col') {
          updated.col = +value;
          if (updated.col === 3) updated.pct = 100;
        } else if (field === 'note') {
          updated.note = value; // keep raw text (do not trim on each keystroke, or spaces get eaten)
        }
        // Keep completion date in sync with Done status
        if (updated.col === 3 && !updated.completedAt) {
          updated.completedAt = new Date().toISOString().split('T')[0];
        } else if (updated.col !== 3) {
          updated.completedAt = undefined;
        }
        updated.updatedAt = new Date().toISOString();
        return updated;
      }
      return t;
    }));
    // Log meaningful transitions (status changes / completion); skip noisy per-keystroke pct edits.
    if (before) {
      if (field === 'col') {
        const toCol = +value;
        if (toCol !== before.col) {
          logActivity(id, before.title, toCol === 3 ? 'done' : 'status', `${COLS[before.col]} → ${COLS[toCol]}`);
        }
      } else if (field === 'pct') {
        const p = Math.min(100, Math.max(0, +value || 0));
        if (p === 100 && before.col !== 3) {
          logActivity(id, before.title, 'done', `${COLS[before.col]} → ${COLS[3]} (100%)`);
        }
      }
    }
  };

  // Subtype Customization
  const handleAddSubtype = () => {
    const val = newSubtypeInput.trim();
    if (!val) return;
    const scope = taskForm.scope;
    if (!subtypes[scope].includes(val)) {
      setSubtypes(prev => ({
        ...prev,
        [scope]: [...prev[scope], val]
      }));
    }
    setTaskForm(prev => ({ ...prev, subtype: val }));
    setNewSubtypeInput('');
    setIsAddingSubtype(false);
  };

  const handleRemoveSubtype = (scope: 'ae' | 'si' | 'pd' | 'va' | 'pr', val: string) => {
    setSubtypes(prev => ({
      ...prev,
      [scope]: prev[scope].filter(s => s !== val)
    }));
    if (taskForm.scope === scope && taskForm.subtype === val) {
      setTaskForm(prev => ({ ...prev, subtype: '' }));
    }
  };

  // Shifting Gantt Window
  const handleShiftTimeline = (days: number) => {
    if (days === 0) {
      const d = new Date();
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() - 14);
      setGanttStart(d);
    } else {
      setGanttStart(prev => {
        const d = new Date(prev);
        d.setDate(prev.getDate() + days);
        return d;
      });
    }
  };

  // Excel Exports
  const handleExportTimelineExcel = () => {
    const rows = tasks.filter(t => t.start || t.deadline).map(t => ({
      'Scope': SCOPES[t.scope],
      'Task': t.title,
      'Assignee': memberName(t.assignee),
      'Sub-type': t.subtype || '',
      'Priority': t.priority.toUpperCase(),
      'Start date': t.start || '',
      'Deadline': t.deadline || '',
      'Status': COLS[t.col],
      '% Complete': t.pct || 0
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [{ wch: 22 }, { wch: 42 }, { wch: 10 }, { wch: 22 }, { wch: 9 }, { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Master Timeline');
    XLSX.writeFile(wb, `IMPACT_Timeline_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Date Label Helper for Kanban
  // --- Weekly filter helpers ---
  // Weeks are month-bounded day chunks: W1 = days 1–7, W2 = 8–14, W3 = 15–21,
  // W4 = 22–28, W5 = 29–end. A week never spans into the next month.
  // Format: "Week {n} - dd/mm/yy - dd/mm/yy"
  const getWeekInfo = (dateStr: string) => {
    const pad = (n: number) => String(n).padStart(2, '0');
    const d = new Date(dateStr); d.setHours(0, 0, 0, 0);
    const year = d.getFullYear(); const month = d.getMonth(); const day = d.getDate();
    const weekOfMonth = Math.ceil(day / 7); // 1..5
    const startDay = (weekOfMonth - 1) * 7 + 1;
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const endDay = Math.min(weekOfMonth * 7, daysInMonth);
    const start = new Date(year, month, startDay);
    const end = new Date(year, month, endDay);
    const key = `${year}-${pad(month + 1)}-${pad(startDay)}`;
    const label = `Week ${weekOfMonth}`;
    const fmt = (dt: Date) => `${pad(dt.getDate())}/${pad(dt.getMonth() + 1)}/${String(dt.getFullYear()).slice(-2)}`;
    const range = `${fmt(start)} - ${fmt(end)}`;
    const full = `${label} - ${range}`;
    return { key, start, end, weekOfMonth, label, range, full };
  };

  // A task belongs to the selected week based on its Start Date
  const inSelectedWeek = (t: Task) => {
    if (weeklyFilter === 'all') return true;
    if (!t.start) return false;
    return getWeekInfo(t.start).key === weeklyFilter;
  };

  // Distinct weeks derived from task Start Dates (+ current week), newest first
  const getWeekOptions = () => {
    const map = new Map<string, ReturnType<typeof getWeekInfo>>();
    tasks.forEach(t => {
      if (t.start) { const wi = getWeekInfo(t.start); if (!map.has(wi.key)) map.set(wi.key, wi); }
    });
    const cur = getWeekInfo(new Date().toISOString().split('T')[0]);
    if (!map.has(cur.key)) map.set(cur.key, cur);
    return [...map.values()].sort((a, b) => b.key.localeCompare(a.key));
  };

  const shiftWeek = (delta: number) => {
    const pad = (n: number) => String(n).padStart(2, '0');
    const baseISO = weeklyFilter === 'all' ? getWeekInfo(new Date().toISOString().split('T')[0]).key : weeklyFilter;
    const wi = getWeekInfo(baseISO);
    // step into the adjacent chunk (handles month boundaries automatically)
    const target = new Date(delta > 0 ? wi.end : wi.start);
    target.setDate(target.getDate() + (delta > 0 ? 1 : -1));
    setWeeklyFilter(getWeekInfo(`${target.getFullYear()}-${pad(target.getMonth() + 1)}-${pad(target.getDate())}`).key);
  };

  // Unified deadline classification used across Kanban, Gantt, and the At-Risk panel.
  // URGENT = past deadline · NEAR = due within 7 days · normal = further out.
  const getDeadlineStatus = (deadline: string) => {
    if (!deadline) return null;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const dd = new Date(deadline); dd.setHours(0, 0, 0, 0);
    const diff = Math.round((dd.getTime() - today.getTime()) / 864e5);
    if (diff < 0) return { kind: 'urgent' as const, label: 'URGENT', diff };
    if (diff <= 7) return { kind: 'near' as const, label: 'NEAR', diff };
    return { kind: 'normal' as const, label: '', diff };
  };

  const getDeadlineBadge = (deadline: string, col: number) => {
    if (!deadline || col === 3 || col === 4) return null;
    const st = getDeadlineStatus(deadline);
    if (!st) return null;
    const dd = new Date(deadline); dd.setHours(0, 0, 0, 0);
    const fmt = dd.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' });

    if (st.kind === 'urgent') {
      return <div className="deadline-badge badge status-badge-danger mt-2">⚠ URGENT · {fmt}</div>;
    }
    if (st.kind === 'near') {
      return <div className="deadline-badge badge status-badge-warning mt-2">⏳ NEAR · {fmt}</div>;
    }
    return <div className="mt-2 text-[10px] text-slate-400">📅 {fmt}</div>;
  };

  // Weekly Date Range Calc
  const getWeeklyDateRange = () => {
    const today = new Date();
    const day = today.getDay();
    const mon = new Date(today);
    mon.setDate(today.getDate() - (day === 0 ? 6 : day - 1));
    const sun = new Date(mon);
    sun.setDate(mon.getDate() + 6);
    return { mon, sun };
  };

  const handlePrintDaily = () => {
    const target = document.getElementById('pg-daily');
    if (target) {
      document.querySelectorAll('.page').forEach(p => p.classList.remove('print-target'));
      target.classList.add('print-target');
      window.print();
    }
  };

  const handlePrintWeekly = () => {
    const target = document.getElementById('pg-weekly');
    if (target) {
      document.querySelectorAll('.page').forEach(p => p.classList.remove('print-target'));
      target.classList.add('print-target');
      window.print();
    }
  };

  // Short date helper for compact cards (dd/mm)
  const fmtShort = (iso?: string) => iso ? new Date(iso).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit' }) : '';

  // Shared task-card renderer used by both the Kanban board and the History Recorded Task box.
  // Compact by default (Title · Priority · Assignee · Due · Status); click to reveal full details.
  const renderTaskCard = (task: Task, colIdx: number) => {
    const expanded = expandedTaskId === task.id;
    const scopeColor = SCOPE_CONFIGS[task.scope].color;
    let overdue = false;
    if (task.deadline && colIdx !== 3 && colIdx !== 4) {
      const d = new Date(task.deadline); d.setHours(0, 0, 0, 0);
      const now = new Date(); now.setHours(0, 0, 0, 0);
      overdue = d < now;
    }
    return (
      <div
        className={`task-card task-card-v2 ${expanded ? 'expanded' : ''} animate-fade-in`}
        key={task.id}
        draggable={isEditMode}
        onDragStart={() => isEditMode && setDraggedTaskId(task.id)}
        onDragEnd={() => setDraggedTaskId(null)}
        style={{ borderLeftColor: scopeColor }}
        onClick={() => setExpandedTaskId(expanded ? null : task.id)}
      >
        <div className="tc-actions edit-only" onClick={e => e.stopPropagation()}>
          <button className="flex items-center gap-1 text-[9px]" onClick={() => handleOpenEditModal(task)}>
            <Edit2 size={8} /> Edit
          </button>
          <button className="btn-r flex items-center gap-1 text-[9px] px-2 py-0.5" onClick={() => handleDeleteTask(task.id)}>
            <Trash2 size={8} /> Del
          </button>
        </div>

        {/* Row 1: title + priority */}
        <div className="tc-top">
          <span className="tc-title font-semibold">{task.title}</span>
          <span className={`pri-badge pri-${task.priority}`}>{task.priority}</span>
        </div>

        {/* Row 2: assignee + due date */}
        <div className="tc-row2">
          <span className="av-tag">
            <span className="av" style={{ background: memberColor(task.assignee), color: '#fff' }}>{memberInitial(task.assignee)}</span>
            {memberName(task.assignee)}
          </span>
          {task.deadline && (
            <span className={`tc-due ${overdue ? 'overdue' : ''}`}>
              <Calendar size={10} /> {fmtShort(task.deadline)}{overdue ? ' ⚠' : ''}
            </span>
          )}
        </div>

        {/* Row 3: status + progress + expand toggle */}
        <div className="tc-row3">
          <span className="stage-badge" style={{ background: `${COL_COLORS[colIdx]}22`, color: COL_COLORS[colIdx], borderColor: `${COL_COLORS[colIdx]}55` }}>
            {COLS[colIdx]}
          </span>
          {colIdx !== 4 && <span className="tc-mini-pct">{task.pct}%</span>}
          <button className="tc-expand-btn" onClick={e => { e.stopPropagation(); setExpandedTaskId(expanded ? null : task.id); }} title={expanded ? 'Thu gọn' : 'Xem chi tiết'}>
            {expanded ? '▲' : '▼'}
          </button>
        </div>

        {/* Expanded details */}
        {expanded && (
          <div className="tc-details" onClick={e => e.stopPropagation()}>
            <div className="tcd-row"><span className="tcd-k">Scope</span><span className={`badge s-${task.scope}`}>{SCOPES[task.scope]}</span></div>
            {task.subtype && <div className="tcd-row"><span className="tcd-k">Sub-type</span><span>{task.subtype}</span></div>}
            {colIdx !== 4 && (
              <div className="tc-pct">
                <div className="pct-bar"><div className="pct-fill" style={{ width: `${task.pct}%` }}></div></div>
                <span className="pct-lbl font-semibold text-slate-400">{task.pct}%</span>
              </div>
            )}
            <div className="tcd-row"><span className="tcd-k">Bắt đầu</span><span>{task.start || '—'}</span></div>
            <div className="tcd-row"><span className="tcd-k">Deadline</span><span className={overdue ? 'tt-overdue' : ''}>{task.deadline || '—'}</span></div>
            {colIdx === 3 && task.completedAt && <div className="tcd-row"><span className="tcd-k">Hoàn thành</span><span className="tc-done-date">✓ {new Date(task.completedAt).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' })}</span></div>}
            {task.link && (
              <a className="tc-link" href={task.link} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()}>
                <LinkIcon size={10} />
                <span className="tc-link-text">{task.link.replace(/^https?:\/\/(www\.)?/, '')}</span>
              </a>
            )}
            {task.desc && <div className="tcd-desc">{task.desc}</div>}
            {task.note && <div className="tcd-note">"{task.note}"</div>}
          </div>
        )}
      </div>
    );
  };

  // Reusable, calm 2-tone combo chart (Target columns + Actual line)
  const comboChart = (opts: { labels: string[]; targets: number[]; actuals: number[]; stacks?: { value: number; color: string }[][]; tip?: (i: number) => React.ReactNode }) => {
    const { labels, targets, actuals, stacks, tip } = opts;
    const fmt = (n: number) => n >= 1e9 ? (n / 1e9).toFixed(n % 1e9 === 0 ? 0 : 1) + ' tỷ' : n >= 1e6 ? (n / 1e6).toFixed(0) + ' tr' : new Intl.NumberFormat('vi-VN').format(Math.round(n || 0));
    const vnd = (n: number) => new Intl.NumberFormat('vi-VN').format(Math.round(n || 0)) + ' ₫';
    const n = labels.length;
    const rawMax = Math.max(1, ...targets, ...actuals);
    const pow = Math.pow(10, Math.floor(Math.log10(rawMax)));
    const yMax = Math.ceil(rawMax / pow) * pow || 1;
    const W = 720, H = 250, padL = 60, padR = 16, padT = 14, padB = 34;
    const plotW = W - padL - padR, plotH = H - padT - padB;
    const bandW = plotW / n;
    const colW = Math.max(10, Math.min(n > 6 ? 46 : 96, bandW * (n > 6 ? 0.5 : 0.62)));
    const yfn = (v: number) => padT + plotH - (v / yMax) * plotH;
    const xc = (i: number) => padL + bandW * i + bandW / 2;
    const linePts = actuals.map((v, i) => `${xc(i)},${yfn(v)}`).join(' ');
    const ticks = [0, 1, 2, 3].map(k => yMax * k / 3);
    return (
      <div className="kpi-combo-wrap">
        <svg className="kpi-combo-svg" viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="xMidYMid meet" onMouseLeave={() => setKpiHoverYear(null)}>
          <defs>
            <linearGradient id="kpiBar" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#818cf8" /><stop offset="100%" stopColor="#6366f1" />
            </linearGradient>
          </defs>
          {ticks.map((v, k) => (
            <g key={k}>
              <line x1={padL} y1={yfn(v)} x2={W - padR} y2={yfn(v)} stroke="var(--panel-border)" strokeWidth={1} />
              <text x={padL - 8} y={yfn(v) + 4} textAnchor="end" className="kpi-combo-axis">{fmt(v)}</text>
            </g>
          ))}
          {kpiHoverYear != null && kpiHoverYear < n && (
            <rect x={padL + bandW * kpiHoverYear} y={padT} width={bandW} height={plotH} fill="var(--color-primary)" opacity={0.07} />
          )}
          {targets.map((v, i) => {
            if (stacks && stacks[i] && stacks[i].length) {
              let cum = 0;
              return (
                <g key={i}>
                  {stacks[i].map((seg, j) => {
                    const yTop = yfn(cum + seg.value);
                    const h = (seg.value / yMax) * plotH;
                    cum += seg.value;
                    return <rect key={j} x={xc(i) - colW / 2} y={yTop} width={colW} height={Math.max(0, h)} fill={seg.color} rx={j === stacks[i].length - 1 ? 3 : 0}><title>{labels[i]}</title></rect>;
                  })}
                </g>
              );
            }
            return <rect key={i} x={xc(i) - colW / 2} y={yfn(v)} width={colW} height={Math.max(0, plotH - (yfn(v) - padT))} fill="url(#kpiBar)" rx={3} />;
          })}
          <polyline points={linePts} fill="none" stroke="var(--text-main)" strokeWidth={1.4} className="kpi-combo-line" />
          {actuals.map((v, i) => <circle key={i} cx={xc(i)} cy={yfn(v)} r={3} fill="var(--panel-bg)" stroke="var(--text-main)" strokeWidth={1.4} />)}
          {labels.map((lb, i) => (
            <g key={i}>
              <text x={xc(i)} y={H - padB + 22} textAnchor="middle" className="kpi-combo-xlabel">{lb}</text>
              <rect x={padL + bandW * i} y={padT} width={bandW} height={plotH} fill="transparent" onMouseEnter={() => setKpiHoverYear(i)} />
            </g>
          ))}
        </svg>
        {kpiHoverYear != null && kpiHoverYear < n && (
          <div className="kpi-combo-tip" style={{ left: `${xc(kpiHoverYear) / W * 100}%`, top: `${padT / H * 100}%` }}>
            <div className="kpi-tip-year">{labels[kpiHoverYear]}</div>
            {tip && tip(kpiHoverYear)}
            <div className="kpi-tip-row tt"><span>🎯 Target</span><b>{vnd(targets[kpiHoverYear])}</b></div>
            <div className="kpi-tip-row tt"><span>📈 Actual</span><b>{vnd(actuals[kpiHoverYear])}</b></div>
            <div className="kpi-tip-row tt"><span>Đạt</span><b className={targets[kpiHoverYear] > 0 && actuals[kpiHoverYear] / targets[kpiHoverYear] >= 1 ? 'up' : targets[kpiHoverYear] > 0 && actuals[kpiHoverYear] / targets[kpiHoverYear] >= 0.7 ? 'mid' : 'down'}>{targets[kpiHoverYear] > 0 ? Math.round(actuals[kpiHoverYear] / targets[kpiHoverYear] * 100) : 0}%</b></div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className={`min-h-screen pb-12 ${isEditMode ? '' : 'read-only'}`}>
      {/* Navigation */}
      <nav className="no-print">
        <div className="nav-logo">IMPACT TEAM</div>
        <div className={`nav-tab ${activeTab === 'dash' ? 'active' : ''}`} onClick={() => setActiveTab('dash')}>Dashboard</div>
        <div className={`nav-tab ${activeTab === 'kanban' ? 'active' : ''}`} onClick={() => setActiveTab('kanban')}>Kanban</div>
        <div className={`nav-tab ${activeTab === 'tracker' ? 'active' : ''}`} onClick={() => setActiveTab('tracker')}>Task Tracker</div>
        <div className={`nav-tab ${activeTab === 'timeline' ? 'active' : ''}`} onClick={() => setActiveTab('timeline')}>Gantt Timeline</div>
        <div className={`nav-tab ${activeTab === 'daily' ? 'active' : ''}`} onClick={() => setActiveTab('daily')}>Daily Report</div>
        <div className={`nav-tab ${activeTab === 'weekly' ? 'active' : ''}`} onClick={() => setActiveTab('weekly')}>Weekly Report</div>
        <div className={`nav-tab ${activeTab === 'kpis' ? 'active' : ''}`} onClick={() => setActiveTab('kpis')}>Target KPIs</div>
        <div className="nav-right">
          <button 
            type="button" 
            onClick={() => setDarkMode(!darkMode)} 
            className="p-1.5 rounded-lg border border-white/10 hover:bg-white/10 text-slate-300 transition-colors mr-2 flex items-center justify-center"
            title={darkMode ? 'Switch to Light Mode' : 'Switch to Dark Mode'}
            style={{ width: '32px', height: '32px', padding: 0 }}
          >
            {darkMode ? <Sun size={14} className="text-yellow-400" /> : <Moon size={14} />}
          </button>

          {isEditMode ? (
            <button className="lock-pill unlocked" onClick={lockToView} title="Khoá lại (chuyển sang Chỉ xem)">
              <Unlock size={13} /> Đang chỉnh sửa
            </button>
          ) : (
            <button className="lock-pill locked" onClick={enterEditMode} title="Nhập mật khẩu để chỉnh sửa">
              <Lock size={13} /> Chỉ xem
            </button>
          )}

          <button className="btn-p edit-only" onClick={() => handleOpenAddModal(0)}>
            <Plus size={14} /> New Task
          </button>
        </div>
      </nav>

      {/* Pages Container */}
      <main className="px-6 mt-6">
        {/* --- 1. DASHBOARD --- */}
        {activeTab === 'dash' && (
          <div className="page active">
            <div className="dash-grid">
              <div className="dash-card">
                <div className="dc-val text-indigo-400">{tasks.length}</div>
                <div className="dc-lbl">Total Tasks</div>
                <div className="dc-sub text-slate-400">
                  <Clock size={11} className="inline text-slate-400" /> {tasks.filter(t => t.col === 1).length} active · {tasks.filter(t => t.col === 3).length} completed
                </div>
              </div>
              <div className="dash-card">
                <div className="dc-val text-emerald-400">
                  {tasks.length ? Math.round(tasks.filter(t => t.col === 3).length / tasks.length * 100) : 0}%
                </div>
                <div className="dc-lbl">Overall Completion</div>
                <div className="dc-sub text-slate-400">
                  <Percent size={11} className="inline" /> Avg progress: {tasks.filter(t => t.col !== 3 && t.col !== 4).length ? Math.round(tasks.filter(t => t.col !== 3 && t.col !== 4).reduce((sum, t) => sum + (t.pct || 0), 0) / tasks.filter(t => t.col !== 3 && t.col !== 4).length) : 0}%
                </div>
              </div>
              <div className="dash-card">
                <div className="dc-val text-rose-500">
                  {tasks.filter(t => {
                    if (!t.deadline || t.col === 3 || t.col === 4) return false;
                    const d = new Date(t.deadline); d.setHours(0,0,0,0);
                    return d < new Date();
                  }).length}
                </div>
                <div className="dc-lbl">Overdue Tasks</div>
                <div className="dc-sub text-slate-400">
                  <AlertCircle size={11} className="inline text-rose-400" /> Urgent focus required
                </div>
              </div>
              <div className="dash-card">
                <div className="dc-val text-amber-500">
                  {tasks.filter(t => t.priority === 'high' && t.col !== 3 && t.col !== 4).length}
                </div>
                <div className="dc-lbl">High Priority Pending</div>
                <div className="dc-sub text-slate-400">
                  <TrendingUp size={11} className="inline text-amber-400" /> Across all projects
                </div>
              </div>
            </div>

            {/* Revenue KPI · Target vs Actual — live combo chart (Year / Month views) */}
            {(() => {
              const fmtC = (n: number) => n >= 1e9 ? (n / 1e9).toFixed(n % 1e9 === 0 ? 0 : 1) + ' tỷ' : n >= 1e6 ? (n / 1e6).toFixed(0) + ' tr' : new Intl.NumberFormat('vi-VN').format(Math.round(n || 0));
              // YEAR view data
              const yData = KPI_YEARS.map(y => {
                const yd = kpis.years[y];
                const target = yd ? yd.channels.reduce((s, c) => s + (yd.totalTarget || 0) * (c.allocation || 0) / 100, 0) : 0;
                const actual = yd ? yd.channels.reduce((s, c) => s + (c.actual ? c.actual.reduce((a, b) => a + (b || 0), 0) : 0), 0) : 0;
                return { y, target, actual, channels: yd?.channels || [], totalTarget: yd?.totalTarget || 0 };
              });
              // MONTH view data (selected dashYear): monthly target & actual across 12 months
              const md = kpis.years[dashYear];
              const monthTargets = Array.from({ length: 12 }, (_, m) => (md?.channels || []).reduce((s, c) => s + (md!.totalTarget || 0) * (c.allocation || 0) / 100 * c.curve[Math.floor(m / 3)] / 3, 0));
              const monthActuals = Array.from({ length: 12 }, (_, m) => (md?.channels || []).reduce((s, c) => s + ((c.actual && c.actual[m]) || 0), 0));
              const monthLbls = ['T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'T9', 'T10', 'T11', 'T12'];

              return (
                <div className="panel kpi-combo-panel">
                  <h3>
                    <span>Revenue KPI · Target vs Actual</span>
                    <div className="kpi-combo-ctrl">
                      <div className="kpi-combo-toggle">
                        <button className={comboView === 'year' ? 'active' : ''} onClick={() => { setComboView('year'); setKpiHoverYear(null); }}>Theo năm</button>
                        <button className={comboView === 'month' ? 'active' : ''} onClick={() => { setComboView('month'); setKpiHoverYear(null); }}>Theo tháng</button>
                      </div>
                      {comboView === 'month' && (
                        <div className="kpi-combo-years">
                          {KPI_YEARS.map(y => <button key={y} className={dashYear === y ? 'active' : ''} onClick={() => { setDashYear(y); setKpiHoverYear(null); }}>{y}</button>)}
                        </div>
                      )}
                    </div>
                  </h3>
                  {comboView === 'year'
                    ? comboChart({
                        labels: KPI_YEARS, targets: yData.map(d => d.target), actuals: yData.map(d => d.actual),
                        stacks: yData.map(d => d.channels.map((c, j) => ({ value: d.totalTarget * (c.allocation || 0) / 100, color: CH_TONES[j % CH_TONES.length] }))),
                        tip: (i) => (
                          <div className="kpi-tip-rows">
                            {yData[i].channels.map((c, j) => (
                              <div className="kpi-tip-row" key={j}><span className="kpi-tip-dot" style={{ background: CH_TONES[j % CH_TONES.length] }}></span>{c.channel}<b>{fmtC(yData[i].totalTarget * (c.allocation || 0) / 100)}</b></div>
                            ))}
                            <div className="kpi-tip-sep"></div>
                          </div>
                        ),
                      })
                    : comboChart({ labels: monthLbls, targets: monthTargets, actuals: monthActuals })}
                  <div className="kpi-combo-legend">
                    {comboView === 'year'
                      ? (kpis.years[KPI_YEARS[0]]?.channels || []).map((c, j) => (
                          <span className="kpi-combo-lg" key={c.id}><i style={{ background: CH_TONES[j % CH_TONES.length] }}></i>{c.channel}</span>
                        ))
                      : <span className="kpi-combo-lg"><i style={{ background: '#6366f1' }}></i>Target Revenue</span>}
                    <span className="kpi-combo-lg"><i className="lg-line"></i>Actual Revenue</span>
                  </div>
                </div>
              );
            })()}

            <div className="dash-row-3">
              {/* Scope Health */}
              <div className="panel">
                <h3>Scope Health <span>% completion by scope</span></h3>
                <div className="flex flex-col gap-4 mt-2">
                  {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => {
                    const st = tasks.filter(t => t.scope === sc);
                    if (!st.length) return null;
                    const pct = Math.round(st.reduce((sum, t) => sum + (t.pct || 0), 0) / st.length);
                    const config = SCOPE_CONFIGS[sc];
                    if (!config) return null;
                    
                    const isGood = pct >= 70;
                    const isWarning = pct >= 40 && pct < 70;
                    const statusText = isGood ? 'On Track' : isWarning ? 'At Risk' : 'Off Track';
                    const statusColor = isGood ? 'text-emerald-400' : isWarning ? 'text-amber-400' : 'text-rose-400';
                    const statusBg = isGood ? 'bg-emerald-950/20' : isWarning ? 'bg-amber-950/20' : 'bg-rose-950/20';

                    return (
                      <div className="scope-row" key={sc}>
                        <div className="risk-dot" style={{ background: isGood ? '#10b981' : isWarning ? '#f59e0b' : '#ef4444', color: isGood ? '#10b981' : isWarning ? '#f59e0b' : '#ef4444' }}></div>
                        <div className="scope-name-lbl font-semibold">{config.name}</div>
                        <div className="scope-track">
                          <div className="scope-track-fill" style={{ width: `${pct}%`, background: config.color }}></div>
                        </div>
                        <div className={`scope-pct-lbl ${statusColor}`}>{pct}%</div>
                        <span className={`badge ${statusBg} ${statusColor} ml-2 font-bold`}>{statusText}</span>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Team Velocity */}
              <div className="panel">
                <h3>Team Velocity <span>tasks completed</span></h3>
                <div className="velocity-bar">
                  {(() => {
                    const today = new Date(); today.setHours(0, 0, 0, 0);
                    // Monday of the current week (week starts Monday)
                    const monday = new Date(today);
                    monday.setDate(today.getDate() - ((today.getDay() + 6) % 7));
                    const weeks = [];
                    for (let i = 5; i >= 0; i--) {
                      const start = new Date(monday); start.setDate(monday.getDate() - i * 7);
                      const end = new Date(start); end.setDate(start.getDate() + 7); // exclusive
                      const weekOfMonth = Math.ceil(start.getDate() / 7);
                      const count = tasks.filter(t => {
                        if (t.col !== 3) return false;
                        const cd = t.completedAt || t.deadline; // fallback for older data
                        if (!cd) return false;
                        const d = new Date(cd); d.setHours(0, 0, 0, 0);
                        return d >= start && d < end;
                      }).length;
                      weeks.push({ label: `W${weekOfMonth}/${start.getMonth() + 1}`, count, isNow: i === 0 });
                    }
                    const max = Math.max(...weeks.map(w => w.count), 1);
                    return weeks.map((w, i) => {
                      const h = w.count === 0 ? 6 : Math.max(10, Math.round((w.count / max) * 70));
                      return (
                        <div className="vel-col-wrap" key={i}>
                          <div className="vel-count">{w.count}</div>
                          <div className="vel-col" style={{ height: `${h}px`, background: w.isNow ? 'var(--color-primary)' : 'rgba(99, 102, 241, 0.3)' }}></div>
                          <div className="vel-lbl">{w.label}</div>
                        </div>
                      );
                    });
                  })()}
                </div>
              </div>
            </div>

            <div className="dash-row">
              {/* Member Workload */}
              <div className="panel">
                <h3>Member Workload</h3>
                <div className="flex flex-col">
                  {members.map(mem => {
                    const m = mem.id;
                    const mt = tasks.filter(t => t.assignee === m);
                    const matched = (['ae', 'si', 'pd', 'va', 'pr'] as const).filter(sc => mt.some(t => t.scope === sc));
                    const scArr = matched.length ? matched : (['ae', 'si', 'pd', 'va', 'pr'] as const);
                    return (
                      <div className="member-load" key={m}>
                        <div className="av w-8 h-8 text-xs font-bold" style={{ background: mem.color, color: '#fff' }}>{memberInitial(m)}</div>
                        <div className="ml-info">
                          <div className="ml-name">{mem.name}</div>
                          <div className="ml-role">{mem.role} · {mt.length} tasks</div>
                        </div>
                        <div className="ml-bar-wrap">
                          {scArr.map(sc => {
                            const st = mt.filter(t => t.scope === sc);
                            const pct = st.length ? Math.round(st.reduce((sum, t) => sum + (t.pct || 0), 0) / st.length) : 0;
                            return (
                              <div className="ml-bar-row" key={sc}>
                                <span className="ml-label text-slate-400 font-semibold">{SCOPES[sc].split(' ')[0]}</span>
                                <div className="ml-bar">
                                  <div className="ml-fill" style={{ width: `${pct}%`, background: SCOPE_CONFIGS[sc].color }}></div>
                                </div>
                                <span className="ml-val">{pct}%</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Overdue & At-Risk list — URGENT (overdue) / NEAR (≤7 days to deadline) */}
              <div className="panel">
                {(() => {
                  const today = new Date(); today.setHours(0, 0, 0, 0);
                  const atRisk = tasks
                    .filter(t => t.deadline && t.col !== 3 && t.col !== 4)
                    .map(t => {
                      const d = new Date(t.deadline); d.setHours(0, 0, 0, 0);
                      const diff = Math.round((d.getTime() - today.getTime()) / 864e5);
                      return { t, diff };
                    })
                    .filter(x => x.diff <= 7)            // overdue OR due within a week
                    .sort((a, b) => a.diff - b.diff);    // most overdue first
                  const urgentCount = atRisk.filter(x => x.diff < 0).length;   // past deadline
                  const nearCount = atRisk.length - urgentCount;               // due in ≤7 days

                  return (
                    <>
                      <h3>
                        <span>Overdue &amp; At-Risk Items</span>
                        <span className="risk-counts">
                          <span className="badge status-badge-danger">{urgentCount} URGENT</span>
                          <span className="badge status-badge-warning">{nearCount} NEAR</span>
                        </span>
                      </h3>
                      <div className="overdue-list">
                        {atRisk.length === 0 && (
                          <div className="text-[11px] text-muted">Không có mục nào quá hạn hoặc sắp tới hạn trong 7 ngày.</div>
                        )}
                        {atRisk.map(({ t, diff }) => {
                          const isUrgent = diff < 0;
                          return (
                            <div className="overdue-item" key={t.id}>
                              <span className={`overdue-item-dot ${isUrgent ? 'bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.6)]' : 'bg-amber-500 shadow-[0_0_8px_rgba(245,158,11,0.6)]'}`}></span>
                              <div className="flex-1">
                                <div className="overdue-item-title font-semibold text-white">{t.title}</div>
                                <div className="overdue-item-meta text-[10px] text-slate-400 flex items-center gap-1 flex-wrap">
                                  <span>{SCOPES[t.scope]} · {memberName(t.assignee)}</span>
                                  <span className={`badge ${isUrgent ? 'status-badge-danger' : 'status-badge-warning'}`}>{isUrgent ? 'URGENT' : 'NEAR'}</span>
                                </div>
                              </div>
                              <span className={`badge ${PRIORITY_BADGES[t.priority]}`}>{t.priority}</span>
                            </div>
                          );
                        })}
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>
          </div>
        )}

        {/* --- 2. KANBAN BOARD --- */}
        {activeTab === 'kanban' && (
          <div className="page active">
            <div className="top-bar">
              <h2>Task Board</h2>
              <div className="filters">
                <button className={`fb ${assigneeFilter === 'all' ? 'active' : ''}`} onClick={() => setAssigneeFilter('all')}>All Assignees</button>
                {members.map(mem => (
                  <button
                    key={mem.id}
                    className={`fb fb-member ${assigneeFilter === mem.id ? 'active' : ''}`}
                    onClick={() => setAssigneeFilter(mem.id)}
                  >
                    <span className="fb-dot" style={{ background: mem.color }}></span>
                    {mem.name.charAt(0) + mem.name.slice(1).toLowerCase()}
                  </button>
                ))}
                <button className="fb fb-add edit-only" onClick={openMemberModal} title="Thêm thành viên">
                  <Plus size={12} />
                </button>
              </div>
              <div className="filters pl-3 border-l border-white/10">
                <button className={`fb ${scopeFilter === 'all' ? 'active' : ''}`} onClick={() => setScopeFilter('all')}>All Scopes</button>
                {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => (
                  <button key={sc} className={`fb ${scopeFilter === sc ? 'active' : ''}`} onClick={() => setScopeFilter(sc)}>{SCOPES[sc]}</button>
                ))}
              </div>
              <div className="search-wrap ml-auto">
                <input 
                  placeholder="🔍 Search tasks..." 
                  value={searchQuery} 
                  onChange={e => setSearchQuery(e.target.value)} 
                />
              </div>
            </div>

            {/* Micro Stats */}
            <div className="stats">
              <div className="sc">
                <div className="n">{tasks.length}</div>
                <div className="l">Total</div>
              </div>
              <div className="sc">
                <div className="n text-amber-400">{tasks.filter(t => t.col === 1).length}</div>
                <div className="l">In Progress</div>
              </div>
              <div className="sc">
                <div className="n text-indigo-400">{tasks.filter(t => t.col === 2).length}</div>
                <div className="l">Review</div>
              </div>
              <div className="sc">
                <div className="n text-emerald-400">{tasks.filter(t => t.col === 3).length}</div>
                <div className="l">Completed</div>
              </div>
              <div className="sc">
                <div className="n text-rose-500">{tasks.filter(t => t.col === 4).length}</div>
                <div className="l">Failed</div>
              </div>
              <div className="sc">
                <div className="n text-red-400">
                  {tasks.filter(t => {
                    if (!t.deadline || t.col === 3 || t.col === 4) return false;
                    const d = new Date(t.deadline); d.setHours(0,0,0,0);
                    return d < new Date();
                  }).length}
                </div>
                <div className="l">Overdue</div>
              </div>
            </div>

            {/* Board Columns */}
            <div className="board mt-2">
              {COLS.map((colName, colIdx) => {
                // Failed / Reject tasks live in the horizontal lane below the board.
                if (colIdx === 4) return null;
                const filteredColTasks = tasks.filter(t => {
                  const matchesAssignee = assigneeFilter === 'all' || t.assignee === assigneeFilter;
                  const matchesScope = scopeFilter === 'all' || t.scope === scopeFilter;
                  const matchesSearch = !searchQuery || t.title.toLowerCase().includes(searchQuery.toLowerCase());
                  return t.col === colIdx && matchesAssignee && matchesScope && matchesSearch;
                });

                return (
                  <div 
                    className={`col ${colIdx === 4 ? 'col-failed' : ''}`} 
                    key={colIdx}
                    onDragOver={e => {
                      e.preventDefault();
                      e.currentTarget.querySelector('.tasks-area')?.classList.add('dov');
                    }}
                    onDragLeave={e => {
                      e.currentTarget.querySelector('.tasks-area')?.classList.remove('dov');
                    }}
                    onDrop={e => {
                      e.preventDefault();
                      e.currentTarget.querySelector('.tasks-area')?.classList.remove('dov');
                      if (draggedTaskId !== null) {
                        updateTaskField(draggedTaskId, 'col', colIdx);
                        setDraggedTaskId(null);
                      }
                    }}
                  >
                    <div className="col-h">
                      <span className="col-title">
                        <span className="col-dot" style={{ background: COL_COLORS[colIdx], color: COL_COLORS[colIdx] }}></span>
                        {colName}
                      </span>
                      <span className="col-cnt">{filteredColTasks.length}</span>
                    </div>

                    <div className="tasks-area">
                      {filteredColTasks.map(task => renderTaskCard(task, colIdx))}
                    </div>

                    {colIdx !== 4 && (
                      <button className="add-btn hover:text-white edit-only" onClick={() => handleOpenAddModal(colIdx)}>
                        <Plus size={12} /> Add task
                      </button>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Failed / Reject — horizontal lane below the board */}
            {(() => {
              const failedTasks = tasks
                .filter(t => {
                  const matchesAssignee = assigneeFilter === 'all' || t.assignee === assigneeFilter;
                  const matchesScope = scopeFilter === 'all' || t.scope === scopeFilter;
                  const matchesSearch = !searchQuery || t.title.toLowerCase().includes(searchQuery.toLowerCase());
                  return t.col === 4 && matchesAssignee && matchesScope && matchesSearch;
                })
                .sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));

              return (
                <div
                  className="history-box lane-failed mt-4"
                  onDragOver={e => { e.preventDefault(); e.currentTarget.classList.add('dov'); }}
                  onDragLeave={e => e.currentTarget.classList.remove('dov')}
                  onDrop={e => {
                    e.preventDefault();
                    e.currentTarget.classList.remove('dov');
                    if (draggedTaskId !== null) {
                      updateTaskField(draggedTaskId, 'col', 4);
                      setDraggedTaskId(null);
                    }
                  }}
                >
                  <div className="history-h">
                    <span className="history-title">
                      <span className="col-dot" style={{ background: COL_COLORS[4], color: COL_COLORS[4] }}></span>
                      Failed / Reject
                    </span>
                    <span className="history-sub">Task bị huỷ / từ chối — kéo task vào đây để đánh dấu Failed</span>
                    <span className="col-cnt ml-auto">{failedTasks.length}</span>
                  </div>
                  <div className="lane-grid">
                    {failedTasks.length === 0 ? (
                      <div className="history-empty">
                        Chưa có task nào bị Failed / Reject. Kéo task vào đây để đánh dấu.
                      </div>
                    ) : (
                      failedTasks.map(task => renderTaskCard(task, 4))
                    )}
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* --- TASK TRACKER (Master table + Activity Log) --- */}
        {activeTab === 'tracker' && (
          <div className="page active">
            {(() => {
              const ACT_META: Record<Activity['type'], { label: string; color: string }> = {
                created: { label: 'Tạo mới', color: '#818cf8' },
                status: { label: 'Chuyển giai đoạn', color: '#3b82f6' },
                progress: { label: 'Tiến độ', color: '#f59e0b' },
                done: { label: 'Hoàn thành', color: '#10b981' },
                edited: { label: 'Chỉnh sửa', color: '#94a3b8' },
                deleted: { label: 'Đã xoá', color: '#ef4444' },
              };
              const fmtDate = (iso?: string) => iso ? new Date(iso).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '—';
              const timeAgo = (iso?: string) => {
                if (!iso) return '—';
                const diff = Date.now() - new Date(iso).getTime();
                const m = Math.floor(diff / 60000);
                if (m < 1) return 'vừa xong';
                if (m < 60) return `${m} phút trước`;
                const h = Math.floor(m / 60);
                if (h < 24) return `${h} giờ trước`;
                const d = Math.floor(h / 24);
                if (d < 30) return `${d} ngày trước`;
                return fmtDate(iso);
              };

              const filtered = tasks.filter(t => {
                const matchesAssignee = assigneeFilter === 'all' || t.assignee === assigneeFilter;
                const matchesScope = scopeFilter === 'all' || t.scope === scopeFilter;
                const matchesSearch = !searchQuery || t.title.toLowerCase().includes(searchQuery.toLowerCase());
                return matchesAssignee && matchesScope && matchesSearch;
              });
              const sorted = [...filtered].sort((a, b) => {
                const dir = trackerSort.dir === 'asc' ? 1 : -1;
                const f = trackerSort.field;
                let av: any, bv: any;
                if (f === 'assignee') { av = memberName(a.assignee); bv = memberName(b.assignee); }
                else if (f === 'scope') { av = SCOPES[a.scope]; bv = SCOPES[b.scope]; }
                else { av = (a as any)[f] ?? ''; bv = (b as any)[f] ?? ''; }
                if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
                return String(av).localeCompare(String(bv)) * dir;
              });
              const today = new Date(); today.setHours(0, 0, 0, 0);
              const arrow = (f: typeof trackerSort.field) => trackerSort.field === f ? (trackerSort.dir === 'asc' ? ' ▲' : ' ▼') : '';
              const cols: { f: typeof trackerSort.field; label: string }[] = [
                { f: 'title', label: 'Task' },
                { f: 'scope', label: 'Scope' },
                { f: 'assignee', label: 'Người phụ trách' },
                { f: 'col', label: 'Giai đoạn' },
                { f: 'pct', label: '% Done' },
                { f: 'deadline', label: 'Deadline' },
                { f: 'updatedAt', label: 'Cập nhật' },
              ];

              return (
                <>
                  <div className="top-bar">
                    <h2>Task Tracker</h2>
                    <div className="filters">
                      <button className={`fb ${assigneeFilter === 'all' ? 'active' : ''}`} onClick={() => setAssigneeFilter('all')}>All Assignees</button>
                      {members.map(mem => (
                        <button key={mem.id} className={`fb fb-member ${assigneeFilter === mem.id ? 'active' : ''}`} onClick={() => setAssigneeFilter(mem.id)}>
                          <span className="fb-dot" style={{ background: mem.color }}></span>
                          {mem.name.charAt(0) + mem.name.slice(1).toLowerCase()}
                        </button>
                      ))}
                    </div>
                    <div className="filters pl-3 border-l border-white/10">
                      <button className={`fb ${scopeFilter === 'all' ? 'active' : ''}`} onClick={() => setScopeFilter('all')}>All Scopes</button>
                      {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => (
                        <button key={sc} className={`fb ${scopeFilter === sc ? 'active' : ''}`} onClick={() => setScopeFilter(sc)}>{SCOPES[sc]}</button>
                      ))}
                    </div>
                    <div className="search-wrap ml-auto">
                      <input placeholder="🔍 Search tasks..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} />
                    </div>
                  </div>

                  {/* Stage summary chips */}
                  <div className="tracker-stages">
                    {COLS.map((colName, idx) => (
                      <div className="tracker-stage-chip" key={idx}>
                        <span className="col-dot" style={{ background: COL_COLORS[idx], color: COL_COLORS[idx] }}></span>
                        <span className="tsc-label">{colName}</span>
                        <span className="tsc-count">{filtered.filter(t => t.col === idx).length}</span>
                      </div>
                    ))}
                  </div>

                  {/* Master tracker table */}
                  <div className="tracker-table-wrap">
                    <table className="tracker-table">
                      <thead>
                        <tr>
                          {cols.map(c => (
                            <th key={c.f} onClick={() => toggleTrackerSort(c.f)} className="tracker-th">
                              {c.label}{arrow(c.f)}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.length === 0 ? (
                          <tr><td colSpan={7} className="tracker-empty">Không có task nào khớp bộ lọc.</td></tr>
                        ) : sorted.map(t => {
                          const overdue = t.deadline && t.col !== 3 && t.col !== 4 && new Date(t.deadline) < today;
                          return (
                            <tr key={t.id} className="tracker-row" onClick={() => isEditMode && handleOpenEditModal(t)}>
                              <td className="tt-title">{t.title}</td>
                              <td><span className={`badge s-${t.scope}`}>{SCOPES[t.scope]}</span></td>
                              <td>
                                <span className="av-tag">
                                  <span className="av" style={{ background: memberColor(t.assignee), color: '#fff' }}>{memberInitial(t.assignee)}</span>
                                  {memberName(t.assignee)}
                                </span>
                              </td>
                              <td>
                                <span className="stage-badge" style={{ background: `${COL_COLORS[t.col]}22`, color: COL_COLORS[t.col], borderColor: `${COL_COLORS[t.col]}55` }}>
                                  {COLS[t.col]}
                                </span>
                              </td>
                              <td>
                                <div className="tt-pct">
                                  <div className="tt-pct-bar"><div className="tt-pct-fill" style={{ width: `${t.pct}%`, background: COL_COLORS[t.col] }}></div></div>
                                  <span className="tt-pct-lbl">{t.pct}%</span>
                                </div>
                              </td>
                              <td className={overdue ? 'tt-overdue' : ''}>
                                {fmtDate(t.deadline)}{overdue ? ' ⚠' : ''}
                              </td>
                              <td className="tt-updated">{timeAgo(t.updatedAt)}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Activity Log — only entries within the last 3 months are kept/shown */}
                  {(() => {
                    const recentActivity = activity.filter(a => isWithinRetention(a.at));
                    return (
                      <div className="activity-section">
                        <div className="activity-h">
                          <Clock size={14} /> Activity Log
                          <span className="activity-sub">Nhật ký mọi thay đổi — tự động giữ trong 3 tháng gần nhất</span>
                          <span className="col-cnt ml-auto">{recentActivity.length}</span>
                        </div>
                        {recentActivity.length === 0 ? (
                          <div className="tracker-empty">Chưa có hoạt động nào trong 3 tháng gần nhất. Tạo / di chuyển / hoàn thành task để bắt đầu theo dõi.</div>
                        ) : (
                          <div className="activity-list">
                            {recentActivity.slice(0, 60).map((a, i) => (
                              <div className="activity-item" key={`${a.at}-${i}`}>
                                <span className="activity-dot" style={{ background: ACT_META[a.type].color }}></span>
                                <span className="activity-type" style={{ color: ACT_META[a.type].color }}>{ACT_META[a.type].label}</span>
                                <span className="activity-title">{a.title}</span>
                                <span className="activity-detail">{a.detail}</span>
                                <span className="activity-time">{timeAgo(a.at)}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </>
              );
            })()}
          </div>
        )}

        {/* --- TARGET KPIs (Annual Revenue Planning by Distribution Channel) --- */}
        {activeTab === 'kpis' && (
          <div className="page active">
            {(() => {
              const vnd = (n: number) => new Intl.NumberFormat('vi-VN').format(Math.round(n || 0));
              const compact = (n: number) => {
                if (n >= 1e9) return (n / 1e9).toFixed(n % 1e9 === 0 ? 0 : 1) + ' tỷ';
                if (n >= 1e6) return (n / 1e6).toFixed(0) + ' tr';
                return vnd(n);
              };
              const total = kpiYearData.totalTarget || 0;
              const allChannels = kpiYearData.channels;
              const visibleChannels = allChannels.filter(c => !kpiChannelFilter || kpiChannelFilter.includes(c.id));
              const rows = visibleChannels.map(c => {
                const annual = total * (c.allocation || 0) / 100;
                const q = c.curve.map(w => annual * w) as number[];
                return { c, annual, q };
              });
              const allocSum = visibleChannels.reduce((s, c) => s + (c.allocation || 0), 0);
              const sum = (f: (r: typeof rows[number]) => number) => rows.reduce((s, r) => s + f(r), 0);
              const tQ = [0, 1, 2, 3].map(i => sum(r => r.q[i]));
              const priCls = (p: string) => p === 'High' ? 'kpi-pri-high' : p === 'Medium' ? 'kpi-pri-med' : 'kpi-pri-low';
              const allIds = allChannels.map(c => c.id);
              const selIds = kpiChannelFilter ?? allIds;
              const toggleChan = (id: number) => {
                const cur = kpiChannelFilter ?? allIds;
                const next = cur.includes(id) ? cur.filter(x => x !== id) : [...cur, id];
                setKpiChannelFilter(next.length === allIds.length ? null : next);
              };

              return (
                <div className="kpi-wrap">
                  {/* Executive title bar */}
                  <div className="kpi-titlebar">
                    <div className="kpi-yearsel">
                      <span className="kpi-year-lbl">Năm kế hoạch</span>
                      <div className="kpi-year-btns">
                        {KPI_YEARS.map(y => (
                          <button key={y} className={`kpi-year-btn ${kpiYear === y ? 'active' : ''}`} onClick={() => setKpiYear(y)}>{y}</button>
                        ))}
                      </div>
                    </div>
                    <div className="kpi-total-box">
                      <div className="kpi-total-lbl">Total Annual Target (VND) · {kpiYear}</div>
                      {isEditMode ? (
                        <input
                          type="text"
                          inputMode="numeric"
                          className="kpi-total-input"
                          value={vnd(total)}
                          onChange={e => updateKpiTotal(e.target.value.replace(/[^\d]/g, ''))}
                        />
                      ) : (
                        <div className="kpi-total-val">{vnd(total)} ₫</div>
                      )}
                      <div className="kpi-total-sub">{compact(total)} · {kpiYear}</div>
                    </div>
                  </div>

                  {/* SECTION 1 header with channel filter */}
                  <div className="kpi-section-bar">
                    <div className="kpi-section-h2">Phân bổ kế hoạch doanh thu theo quý</div>
                    <div className="kpi-chan-filter">
                      <button className="kpi-chan-btn" onClick={() => setKpiChanMenuOpen(o => !o)}>
                        Kênh: {kpiChannelFilter ? `${selIds.length}/${allIds.length}` : `Tất cả (${allIds.length})`} ▾
                      </button>
                      {kpiChanMenuOpen && (
                        <>
                          <div className="kpi-menu-backdrop" onClick={() => setKpiChanMenuOpen(false)}></div>
                          <div className="kpi-chan-menu">
                            <button className="kpi-chan-all" onClick={() => setKpiChannelFilter(null)}>Chọn tất cả</button>
                            {allChannels.map(c => (
                              <label key={c.id} className="kpi-chan-item">
                                <input type="checkbox" checked={selIds.includes(c.id)} onChange={() => toggleChan(c.id)} />
                                <span>{c.channel}</span>
                              </label>
                            ))}
                          </div>
                        </>
                      )}
                    </div>
                  </div>
                  <div className="kpi-table-wrap">
                    <table className="kpi-table">
                      <thead>
                        <tr>
                          <th className="kpi-sticky-col">Distribution Channel</th>
                          <th>Allocation %</th>
                          <th>Annual Target (VND)</th>
                          <th>Q1</th><th>Q2</th><th>Q3</th><th>Q4</th>
                          <th className="kpi-pri-th">Priority</th>
                          <th className="kpi-notes-col">Notes</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((r, idx) => (
                          <tr key={r.c.id} className={idx % 2 ? 'kpi-alt' : ''}>
                            <td className="kpi-sticky-col kpi-channel">{r.c.channel}</td>
                            <td className="kpi-num">
                              {isEditMode ? (
                                <span className="kpi-alloc-wrap">
                                  <input type="number" className="kpi-cell-input kpi-alloc-input" value={r.c.allocation} onChange={e => updateKpiChannel(r.c.id, 'allocation', e.target.value)} />
                                  <span className="kpi-alloc-pct">%</span>
                                </span>
                              ) : `${r.c.allocation}%`}
                            </td>
                            <td className="kpi-num kpi-money">{vnd(r.annual)}</td>
                            {r.q.map((v, i) => <td key={i} className="kpi-num">{compact(v)}</td>)}
                            <td className="kpi-pri-cell">
                              {isEditMode ? (
                                <select className={`kpi-cell-input kpi-pri-select ${priCls(r.c.priority)}`} value={r.c.priority} onChange={e => updateKpiChannel(r.c.id, 'priority', e.target.value)}>
                                  <option>High</option><option>Medium</option><option>Low</option>
                                </select>
                              ) : <span className={`kpi-pri ${priCls(r.c.priority)}`}>{r.c.priority}</span>}
                            </td>
                            <td className="kpi-notes-col">
                              {isEditMode ? (
                                <input className="kpi-cell-input kpi-notes-input" value={r.c.notes} onChange={e => updateKpiChannel(r.c.id, 'notes', e.target.value)} />
                              ) : <span className="kpi-notes-text">{r.c.notes}</span>}
                            </td>
                          </tr>
                        ))}
                        {rows.length === 0 && (
                          <tr><td className="kpi-sticky-col" colSpan={9} style={{ textAlign: 'center', color: 'var(--text-muted)' }}>Chưa chọn kênh nào — mở bộ lọc "Kênh" để hiển thị.</td></tr>
                        )}
                        <tr className="kpi-total-row">
                          <td className="kpi-sticky-col">TOTAL</td>
                          <td className="kpi-num">{allocSum.toFixed(1)}%</td>
                          <td className="kpi-num kpi-money">{vnd(sum(r => r.annual))}</td>
                          {tQ.map((v, i) => <td key={i} className="kpi-num">{compact(v)}</td>)}
                          <td></td><td></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* SECTION 2 — Monthly Target Allocation */}
                  {(() => {
                    const months = ['Tháng 1', 'Tháng 2', 'Tháng 3', 'Tháng 4', 'Tháng 5', 'Tháng 6', 'Tháng 7', 'Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11', 'Tháng 12'];
                    // Each KPI row: 12 monthly values derived from the channel's quarterly curve (quarter / 3 per month).
                    let mRows = rows.map(r => ({
                      c: r.c,
                      m: Array.from({ length: 12 }, (_, i) => r.annual * r.c.curve[Math.floor(i / 3)] / 3),
                      a: Array.from({ length: 12 }, (_, i) => (r.c.actual && r.c.actual[i]) || 0),
                      total: r.annual,
                    }));
                    const dir = kpiMonthSort.dir === 'asc' ? 1 : -1;
                    mRows = [...mRows].sort((a, b) => {
                      if (kpiMonthSort.col === 'name') return a.c.channel.localeCompare(b.c.channel) * dir;
                      if (kpiMonthSort.col === 'total') return (a.total - b.total) * dir;
                      return (a.m[kpiMonthSort.col as number] - b.m[kpiMonthSort.col as number]) * dir;
                    });
                    const colTotals = Array.from({ length: 12 }, (_, i) => mRows.reduce((s, r) => s + r.m[i], 0));
                    const actTotals = Array.from({ length: 12 }, (_, i) => mRows.reduce((s, r) => s + r.a[i], 0));
                    const grand = mRows.reduce((s, r) => s + r.total, 0);
                    const grandAct = mRows.reduce((s, r) => s + r.a.reduce((x, y) => x + y, 0), 0);
                    const arrow = (col: 'name' | 'total' | number) => kpiMonthSort.col === col ? (kpiMonthSort.dir === 'asc' ? ' ↑' : ' ↓') : '';
                    const actCls = (act: number, tgt: number) => act <= 0 ? '' : act >= tgt ? 'kpi-act-up' : 'kpi-act-down';
                    const achCls = (act: number, tgt: number) => act <= 0 ? 'kpi-ach-none' : act >= tgt ? 'kpi-ach-up' : act >= tgt * 0.7 ? 'kpi-ach-mid' : 'kpi-ach-down';

                    return (
                      <>
                        <div className="kpi-section-bar">
                          <div className="kpi-section-h2">PHÂN BỔ MỤC TIÊU THEO THÁNG</div>
                        </div>
                        <div className="kpi-table-wrap">
                          <table className="kpi-table kpi-month-table">
                            <thead>
                              <tr>
                                <th className="kpi-sticky-col kpi-th-sort" onClick={() => toggleKpiMonthSort('name')}>Hạng mục / KPI{arrow('name')}</th>
                                {months.map((m, i) => (
                                  <React.Fragment key={i}>
                                    <th className="kpi-th-sort" onClick={() => toggleKpiMonthSort(i)}>{m}{arrow(i)}</th>
                                    <th className="kpi-actual-col">{`T${i + 1} Actual`}</th>
                                  </React.Fragment>
                                ))}
                                <th className="kpi-th-sort kpi-total-col" onClick={() => toggleKpiMonthSort('total')}>Tổng KPI{arrow('total')}</th>
                                <th className="kpi-total-col kpi-actual-col">Tổng Actual KPI</th>
                              </tr>
                            </thead>
                            <tbody>
                              {mRows.map((r, idx) => {
                                const actSum = r.a.reduce((x, y) => x + y, 0);
                                return (
                                  <tr key={r.c.id} className={idx % 2 ? 'kpi-alt' : ''}>
                                    <td className="kpi-sticky-col kpi-channel">{r.c.channel}</td>
                                    {r.m.map((v, i) => (
                                      <React.Fragment key={i}>
                                        <td className="kpi-num">{compact(v)}</td>
                                        <td className="kpi-num kpi-actual-col">
                                          {isEditMode ? (
                                            <input type="number" className="kpi-cell-input kpi-actual-input" value={r.a[i] || 0} onChange={e => updateKpiActual(r.c.id, i, e.target.value)} />
                                          ) : <span className={actCls(r.a[i], v)}>{r.a[i] ? compact(r.a[i]) : '—'}</span>}
                                        </td>
                                      </React.Fragment>
                                    ))}
                                    <td className="kpi-num kpi-total-col kpi-money">{compact(r.total)}</td>
                                    <td className="kpi-num kpi-total-col kpi-actual-col">
                                      <div className="kpi-totalcell">
                                        <span className="kpi-money">{actSum ? compact(actSum) : '—'}</span>
                                        <span className={`kpi-ach ${achCls(actSum, r.total)}`}>{r.total > 0 && actSum > 0 ? Math.round(actSum / r.total * 100) + '%' : '—'}</span>
                                      </div>
                                    </td>
                                  </tr>
                                );
                              })}
                              {mRows.length === 0 && (
                                <tr><td className="kpi-sticky-col" colSpan={27} style={{ textAlign: 'center', color: 'var(--text-muted)' }}>Không có hạng mục nào khớp bộ lọc.</td></tr>
                              )}
                              <tr className="kpi-total-row">
                                <td className="kpi-sticky-col">TỔNG</td>
                                {colTotals.map((v, i) => (
                                  <React.Fragment key={i}>
                                    <td className="kpi-num">{compact(v)}</td>
                                    <td className="kpi-num kpi-actual-col">{actTotals[i] ? compact(actTotals[i]) : '—'}</td>
                                  </React.Fragment>
                                ))}
                                <td className="kpi-num kpi-total-col kpi-money">{compact(grand)}</td>
                                <td className="kpi-num kpi-total-col kpi-actual-col kpi-money">{grandAct ? compact(grandAct) : '—'}</td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </>
                    );
                  })()}
                </div>
              );
            })()}
          </div>
        )}

        {/* --- 3. GANTT TIMELINE --- */}
        {activeTab === 'timeline' && (
          <div className="page active">
            <div className="tl-controls">
              <h2>Gantt Timeline</h2>
              <div className="tl-scope-tabs">
                <button className={`tl-tab ${timelineScopeFilter === 'all' ? 'active' : ''}`} onClick={() => setTimelineScopeFilter('all')}>All Scopes</button>
                {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => (
                  <button key={sc} className={`tl-tab ${timelineScopeFilter === sc ? 'active' : ''}`} onClick={() => setTimelineScopeFilter(sc)}>{SCOPES[sc]}</button>
                ))}
              </div>
              <div className="ml-auto flex gap-2">
                <button onClick={() => handleShiftTimeline(-14)}><ChevronLeft size={12} /> 2w</button>
                <button className="btn-p" onClick={() => handleShiftTimeline(0)}>Today</button>
                <button onClick={() => handleShiftTimeline(14)}>2w <ChevronRight size={12} /></button>
                <button className="btn-g" onClick={handleExportTimelineExcel}>
                  <Download size={12} /> Export Excel
                </button>
              </div>
            </div>

            {/* Gantt Render */}
            {(() => {
              const today = new Date(); today.setHours(0,0,0,0);
              const days = [];
              for(let i=0; i<GANTT_DAYS; i++) {
                const d = new Date(ganttStart);
                d.setDate(ganttStart.getDate() + i);
                days.push(d);
              }

              const months: { label: string; count: number }[] = [];
              days.forEach(d => {
                const label = d.toLocaleDateString('vi-VN', { month: 'short', year: 'numeric' });
                if (!months.length || months[months.length - 1].label !== label) {
                  months.push({ label, count: 1 });
                } else {
                  months[months.length - 1].count++;
                }
              });

              const DAY_W = 24;
              const LABEL_W = 220;
              const totalW = LABEL_W + GANTT_DAYS * DAY_W;

              const scopeOrder = ['ae', 'si', 'pd', 'va', 'pr'] as const;
              const activeScopes = timelineScopeFilter === 'all' ? scopeOrder : [timelineScopeFilter];
              const timelineTasks = tasks.filter(t => activeScopes.includes(t.scope) && (t.start || t.deadline));

              return (
                <div className="gantt-container overflow-x-auto">
                  <div style={{ minWidth: `${totalW}px` }}>
                    {/* Header */}
                    <div className="gantt-header" style={{ gridTemplateColumns: `${LABEL_W}px 1fr` }}>
                      <div className="gantt-header-left font-bold text-slate-400">Tasks</div>
                      <div>
                        <div className="gantt-months">
                          {months.map((m, i) => (
                            <div className="gantt-month" style={{ width: `${m.count * DAY_W}px`, flexShrink: 0 }} key={i}>{m.label}</div>
                          ))}
                        </div>
                        <div className="gantt-days">
                          {days.map((d, i) => {
                            const isToday = d.getTime() === today.getTime();
                            const isWknd = d.getDay() === 0 || d.getDay() === 6;
                            return (
                              <div className={`gantt-day ${isToday ? 'today' : ''} ${isWknd ? 'weekend' : ''}`} style={{ width: `${DAY_W}px`, flexShrink: 0 }} key={i}>
                                {d.getDate()}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    </div>

                    {/* Gantt Rows */}
                    <div className="gantt-body">
                      {activeScopes.map(sc => {
                        const st = timelineTasks.filter(t => t.scope === sc);
                        if (!st.length) return null;

                        return (
                          <React.Fragment key={sc}>
                            <div className="gantt-scope-header" style={{ width: `${totalW}px` }}>
                              <span className="w-2 h-2 rounded-full inline-block" style={{ background: SCOPE_CONFIGS[sc].color }}></span>
                              <span className="font-bold text-xs uppercase tracking-wide">
                                {SCOPES[sc]} &nbsp;·&nbsp; {sc === 'pr' ? 'MERCURY' : 'TIFFANY'} &nbsp;·&nbsp; {st.length} tasks
                              </span>
                            </div>

                            {st.map(t => {
                              const s = t.start ? new Date(t.start) : new Date(t.deadline); s.setHours(0,0,0,0);
                              const e = t.deadline ? new Date(t.deadline) : s; e.setHours(0,0,0,0);
                              
                              const offsetDays = Math.round((s.getTime() - ganttStart.getTime()) / 864e5);
                              const spanDays = Math.max(1, Math.round((e.getTime() - s.getTime()) / 864e5) + 1);
                              
                              const leftPx = offsetDays * DAY_W;
                              const widthPx = spanDays * DAY_W - 2;
                              const isVisible = leftPx < GANTT_DAYS * DAY_W && leftPx + widthPx > 0;
                              
                              const clampedLeft = Math.max(0, leftPx);
                              const clampedW = Math.min(widthPx, GANTT_DAYS * DAY_W - clampedLeft);
                              
                              const dStatus = (t.col !== 3 && t.col !== 4) ? getDeadlineStatus(t.deadline) : null;
                              const barColor = t.col === 3 ? 'var(--color-success)' : t.col === 4 ? '#64748b' : dStatus?.kind === 'urgent' ? 'var(--color-danger)' : dStatus?.kind === 'near' ? 'var(--color-warning)' : SCOPE_CONFIGS[sc]?.color || '#ccc';
                              const progressW = Math.round((t.pct || 0) / 100 * clampedW);

                              return (
                                <div className="gantt-row" style={{ gridTemplateColumns: `${LABEL_W}px 1fr`, width: `${totalW}px` }} key={t.id}>
                                  <div className="gantt-label">
                                    <div className="gl-title font-semibold">{t.title}</div>
                                    <div className="gl-sub">
                                      <span className="av" style={{ background: memberColor(t.assignee), color: '#fff' }}>{memberInitial(t.assignee)}</span>
                                      <span>{t.subtype || SCOPES[t.scope]}</span>
                                      {t.col === 3 && <span className="text-emerald-400 font-semibold">✓ Done</span>}
                                      {t.col === 4 && <span className="text-slate-400 font-semibold">✕ Reject</span>}
                                      {dStatus?.kind === 'urgent' && <span className="badge status-badge-danger">⚠ URGENT</span>}
                                      {dStatus?.kind === 'near' && <span className="badge status-badge-warning">⏳ NEAR</span>}
                                    </div>
                                  </div>
                                  
                                  <div className="gantt-chart-area h-11 relative overflow-hidden">
                                    {/* Grid columns background */}
                                    {days.map((_, dayIdx) => (
                                      <div 
                                        className={`gantt-grid-line ${days[dayIdx].getTime() === today.getTime() ? 'today-line' : ''} ${days[dayIdx].getDay() === 0 || days[dayIdx].getDay() === 6 ? 'weekend-line' : ''}`}
                                        style={{ left: `${dayIdx * DAY_W}px` }}
                                        key={dayIdx}
                                      />
                                    ))}

                                    {/* Gantt Bar */}
                                    {isVisible && (
                                      <div 
                                        className="gantt-bar-g absolute top-1/2 -translate-y-1/2" 
                                        style={{ left: `${clampedLeft}px`, width: `${clampedW}px`, background: barColor }}
                                        onClick={() => handleOpenEditModal(t)}
                                      >
                                        <div className="gb-progress" style={{ width: `${progressW}px` }}></div>
                                        <span className="gb-label relative z-10 truncate pl-1 font-bold text-[9px]">
                                          {clampedW > 60 ? t.title : ''}
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </React.Fragment>
                        );
                      })}
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* Timeline Legend */}
            <div className="tl-legend mt-4">
              <div className="text-[11px] font-bold text-slate-400 mr-2 uppercase">Legend:</div>
              {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => (
                <div className="tl-legend-item" key={sc}>
                  <div className="tl-legend-dot" style={{ background: SCOPE_CONFIGS[sc].color }}></div>
                  {SCOPES[sc]}
                </div>
              ))}
              <div className="tl-legend-item"><div className="tl-legend-dot bg-emerald-500"></div>Completed</div>
              <div className="tl-legend-item"><div className="tl-legend-dot bg-rose-500"></div>URGENT (quá hạn)</div>
              <div className="tl-legend-item"><div className="tl-legend-dot bg-amber-500"></div>NEAR (≤7 ngày)</div>
              <div className="tl-legend-item"><div className="tl-legend-dot bg-slate-500"></div>Rejected</div>
              <div className="text-[10px] text-slate-400 italic ml-auto">Click on any bar to edit task directly</div>
            </div>
          </div>
        )}

        {/* --- 4. DAILY PROGRESS REPORT --- */}
        {activeTab === 'daily' && (
          <div className="page active" id="pg-daily">
            <div className="daily-wrap">
              <div className="daily-header">
                <div>
                  <h2>Daily Progress Report — Team IMPACT</h2>
                  <div className="daily-date font-bold text-main" id="daily-date-label" style={{ color: 'var(--color-primary)' }}>
                    {new Date(dailyDate).toLocaleDateString('vi-VN', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' })}
                  </div>
                </div>
                <div className="flex gap-2 items-center no-print">
                  <button className="btn-g" onClick={handlePrintDaily}>
                    <FileText size={12} /> Export PDF
                  </button>
                </div>
              </div>

              {/* Filter header — filter by Start Date */}
              <div className="daily-filter no-print">
                <span className="daily-filter-label">Lọc theo Start Date:</span>
                <div className="daily-filter-toggle">
                  <button
                    className={`fb ${dailyFilterMode === 'date' ? 'active' : ''}`}
                    onClick={() => setDailyFilterMode('date')}
                  >
                    Theo ngày
                  </button>
                  <button
                    className={`fb ${dailyFilterMode === 'all' ? 'active' : ''}`}
                    onClick={() => setDailyFilterMode('all')}
                  >
                    Tất cả
                  </button>
                </div>
                <input
                  type="date"
                  value={dailyDate}
                  disabled={dailyFilterMode === 'all'}
                  onChange={e => { setDailyDate(e.target.value); setDailyFilterMode('date'); }}
                  className="daily-date-picker"
                />
                <button
                  className="fb"
                  onClick={() => { setDailyDate(new Date().toISOString().split('T')[0]); setDailyFilterMode('date'); }}
                >
                  Hôm nay
                </button>
                <span className="daily-filter-hint">
                  {dailyFilterMode === 'date'
                    ? 'Chỉ hiện task có Start Date đúng ngày đã chọn'
                    : 'Đang hiển thị tất cả task'}
                </span>
              </div>

              {(() => {
                const scopesWithTasks = (['ae', 'si', 'pd', 'va', 'pr'] as const).filter(sc =>
                  tasks.some(t => t.scope === sc && (dailyFilterMode === 'all' || t.start === dailyDate))
                );
                if (scopesWithTasks.length === 0) {
                  return (
                    <div className="daily-empty">
                      Không có task nào có Start Date là{' '}
                      <strong>{new Date(dailyDate).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' })}</strong>.
                      Chọn ngày khác hoặc bấm "Tất cả".
                    </div>
                  );
                }
                return null;
              })()}

              {/* Sections for each scope */}
              {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => {
                const st = tasks
                  .filter(t => t.scope === sc && (dailyFilterMode === 'all' || t.start === dailyDate))
                  .sort((a, b) => (a.start || '').localeCompare(b.start || ''));
                if (!st.length) return null;

                return (
                  <div className="daily-section mt-4" key={sc}>
                    <div className="daily-section-h flex items-center gap-2">
                      <span className="w-2 h-2 rounded-full" style={{ background: SCOPE_CONFIGS[sc].color }}></span>
                      <span className="daily-section-title font-bold text-sm">{SCOPES[sc]}</span>
                    </div>
                    <table className="daily-table w-full text-left">
                      <thead>
                        <tr>
                          <th className="p-4" style={{ width: '280px' }}>Task</th>
                          <th className="p-4" style={{ width: '150px' }}>Sub-type</th>
                          <th className="p-4" style={{ width: '120px' }}>Assignee</th>
                          <th className="p-4 text-center" style={{ width: '100px' }}>% Done</th>
                          <th className="p-4" style={{ width: '140px' }}>Status</th>
                          <th className="p-4">Notes</th>
                        </tr>
                      </thead>
                      <tbody>
                        {st.map(t => (
                          <tr key={t.id}>
                            <td className="p-4 font-semibold">{t.title}</td>
                            <td className="p-4 text-muted">{t.subtype || '—'}</td>
                            <td className="p-4 assignee-cell">
                              <div className="flex items-center gap-2">
                                <span className="av" style={{ background: memberColor(t.assignee), color: '#fff' }}>{memberInitial(t.assignee)}</span>
                                <span className="font-semibold assignee-name">{memberName(t.assignee)}</span>
                              </div>
                            </td>
                            <td className="p-4 text-center">
                              <input
                                type="number"
                                className="pct-input"
                                min="0"
                                max="100"
                                value={t.pct}
                                disabled={!isEditMode}
                                onChange={e => updateTaskField(t.id, 'pct', e.target.value)}
                              />
                            </td>
                            <td className="p-4">
                              <select
                                className="status-sel"
                                value={t.col}
                                disabled={!isEditMode}
                                onChange={e => updateTaskField(t.id, 'col', e.target.value)}
                              >
                                {COLS.map((colName, idx) => (
                                  <option value={idx} key={idx}>{colName}</option>
                                ))}
                              </select>
                            </td>
                            <td className="p-4">
                              <textarea
                                className="note-input"
                                rows={2}
                                value={t.note || ''}
                                placeholder="Update progress notes..."
                                disabled={!isEditMode}
                                onChange={e => updateTaskField(t.id, 'note', e.target.value)}
                              />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              })}

              {/* Summary — reflects the active Start Date filter */}
              {(() => {
                const dailyTasks = tasks.filter(t => dailyFilterMode === 'all' || t.start === dailyDate);
                return (
                  <div className="daily-summary mt-6">
                    <h3>Daily Summary Stats {dailyFilterMode === 'date' && <span className="text-muted">· theo Start Date</span>}</h3>
                    <div className="summary-grid">
                      <div className="sum-card">
                        <div className="sn">{dailyTasks.length}</div>
                        <div className="sl text-muted">Total Tasks</div>
                      </div>
                      <div className="sum-card">
                        <div className="sn text-success">{dailyTasks.filter(t => t.col === 3).length}</div>
                        <div className="sl text-muted">Completed</div>
                      </div>
                      <div className="sum-card">
                        <div className="sn text-warning">{dailyTasks.filter(t => t.col === 1).length}</div>
                        <div className="sl text-muted">In Progress</div>
                      </div>
                      <div className="sum-card">
                        <div className="sn text-main" style={{ color: 'var(--color-primary)' }}>
                          {dailyTasks.length ? Math.round(dailyTasks.reduce((sum, t) => sum + t.pct, 0) / dailyTasks.length) : 0}%
                        </div>
                        <div className="sl text-muted">Avg Progress</div>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>
        )}

        {/* --- 5. WEEKLY REPORT --- */}
        {activeTab === 'weekly' && (
          <div className="page active" id="pg-weekly">
            <div className="weekly-wrap">
              <div className="weekly-header">
                <div>
                  <h2>Weekly Report — Team IMPACT</h2>
                  <div className="weekly-to">
                    {weeklyFilter === 'all'
                      ? <>Hiển thị <strong>tất cả các tuần</strong></>
                      : <strong>{getWeekInfo(weeklyFilter).full}</strong>}
                  </div>
                </div>
                <button className="btn-g no-print" onClick={handlePrintWeekly}>
                  <FileText size={12} /> Export PDF
                </button>
              </div>

              {/* Filter header — filter by week (derived from Start Date) */}
              <div className="daily-filter no-print mt-4">
                <span className="daily-filter-label">Lọc theo tuần (Start Date):</span>
                <button className="fb" onClick={() => shiftWeek(-1)} title="Tuần trước">‹</button>
                <select
                  className="daily-date-picker week-select"
                  value={weeklyFilter}
                  onChange={e => setWeeklyFilter(e.target.value)}
                >
                  <option value="all">Tất cả các tuần</option>
                  {getWeekOptions().map(w => (
                    <option value={w.key} key={w.key}>{w.full}</option>
                  ))}
                </select>
                <button className="fb" onClick={() => shiftWeek(1)} title="Tuần sau">›</button>
                <button
                  className="fb"
                  onClick={() => setWeeklyFilter(getWeekInfo(new Date().toISOString().split('T')[0]).key)}
                >
                  Tuần này
                </button>
                <button
                  className={`fb ${weeklyFilter === 'all' ? 'active' : ''}`}
                  onClick={() => setWeeklyFilter('all')}
                >
                  Tất cả
                </button>
                <span className="daily-filter-hint">
                  {weeklyFilter === 'all'
                    ? 'Đang hiển thị mọi task'
                    : 'Chỉ tổng hợp task có Start Date trong tuần đã chọn'}
                </span>
              </div>

              {weeklyFilter !== 'all' && !tasks.some(t => inSelectedWeek(t)) && (
                <div className="daily-empty">
                  Không có task nào có Start Date trong tuần <strong>{getWeekInfo(weeklyFilter).label} ({getWeekInfo(weeklyFilter).range})</strong>.
                  Chọn tuần khác hoặc bấm "Tất cả".
                </div>
              )}

              {/* Revenue KPI — current (realtime) year + 12-month trend, live from Target KPIs */}
              {(() => {
                const fmt = (n: number) => n >= 1e9 ? (n / 1e9).toFixed(n % 1e9 === 0 ? 0 : 1) + ' tỷ' : n >= 1e6 ? (n / 1e6).toFixed(0) + ' tr' : new Intl.NumberFormat('vi-VN').format(Math.round(n || 0));
                const curY = String(new Date().getFullYear());
                const year = kpis.years[curY] ? curY : (KPI_YEARS.find(y => kpis.years[y]) || KPI_YEARS[0]);
                const yd = kpis.years[year];
                const target = yd ? yd.channels.reduce((s, c) => s + (yd.totalTarget || 0) * (c.allocation || 0) / 100, 0) : 0;
                const actual = yd ? yd.channels.reduce((s, c) => s + (c.actual ? c.actual.reduce((a, b) => a + (b || 0), 0) : 0), 0) : 0;
                const pct = target > 0 ? Math.round(actual / target * 100) : 0;
                const cls = pct >= 100 ? 'up' : pct >= 70 ? 'mid' : 'down';
                const monthTargets = Array.from({ length: 12 }, (_, m) => (yd?.channels || []).reduce((s, c) => s + (yd!.totalTarget || 0) * (c.allocation || 0) / 100 * c.curve[Math.floor(m / 3)] / 3, 0));
                const monthActuals = Array.from({ length: 12 }, (_, m) => (yd?.channels || []).reduce((s, c) => s + ((c.actual && c.actual[m]) || 0), 0));
                const monthLbls = ['T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'T9', 'T10', 'T11', 'T12'];
                return (
                  <div className="panel kpi-live mt-4">
                    <h3>
                      <span>Revenue KPI · Target vs Actual</span>
                      <span className="kpi-live-badge">Năm {year} · realtime</span>
                    </h3>
                    <div className="kpi-live-summary">
                      <div className="kls-item"><div className="kls-lbl">🎯 Target</div><div className="kls-val">{fmt(target)}</div></div>
                      <div className="kls-item"><div className="kls-lbl">📈 Actual</div><div className="kls-val">{fmt(actual)}</div></div>
                      <div className="kls-item"><div className="kls-lbl">Đạt KPI</div><div className={`kls-val ${cls}`}>{pct}%</div></div>
                      <div className="kls-track"><div className={`kls-fill ${cls}`} style={{ width: `${Math.min(100, pct)}%` }}></div></div>
                    </div>
                    <div className="kpi-live-monthh">Biến động doanh thu theo 12 tháng — {year}</div>
                    {comboChart({ labels: monthLbls, targets: monthTargets, actuals: monthActuals })}
                    <div className="kpi-combo-legend">
                      <span className="kpi-combo-lg"><i style={{ background: '#6366f1' }}></i>Target Revenue</span>
                      <span className="kpi-combo-lg"><i className="lg-line"></i>Actual Revenue</span>
                    </div>
                  </div>
                );
              })()}

              {/* Weekly content - KPIs by Scope */}
              <div className="scope-kpi-grid mt-4">
                {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => {
                  const st = tasks.filter(t => t.scope === sc && inSelectedWeek(t));
                  if (!st.length) return null;

                  const pct = Math.round(st.reduce((sum, t) => sum + (t.pct || 0), 0) / st.length);
                  const ov = st.filter(t => {
                    if (!t.deadline || t.col === 3 || t.col === 4) return false;
                    const d = new Date(t.deadline); d.setHours(0,0,0,0);
                    return d < new Date();
                  }).length;

                  const isGood = pct >= 70;
                  const isWarning = pct >= 40 && pct < 70;
                  const statusLbl = isGood ? 'On Track' : isWarning ? 'At Risk' : 'Off Track';

                  return (
                    <div className="kpi-card" key={sc}>
                      <div className="kpi-card-h">
                        <span className="scope-name border-l-4 pl-2 font-bold" style={{ borderColor: SCOPE_CONFIGS[sc].color }}>SOW {SCOPES[sc]}</span>
                        <span className={`risk-badge ${isGood ? 'risk-on' : isWarning ? 'risk-at' : 'risk-off'}`}>{statusLbl}</span>
                      </div>
                      <div className="kpi-metrics">
                        <div className="kpi-m"><div className="kv">{st.length}</div><div className="kl">Total</div></div>
                        <div className="kpi-m"><div className="kv text-success">{st.filter(t => t.col === 3).length}</div><div className="kl">Done</div></div>
                        <div className="kpi-m"><div className="kv text-danger">{ov}</div><div className="kl">Overdue</div></div>
                      </div>
                      <div className="scope-bar-wrap">
                        <div className="scope-bar">
                          <div className="scope-bar-fill" style={{ width: `${pct}%`, background: SCOPE_CONFIGS[sc].color }}></div>
                        </div>
                        <span className="scope-bar-pct font-bold">{pct}%</span>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Workload */}
              <div className="workload-section">
                <h3>Weekly Workload Summary</h3>
                {members.map(mem => {
                  const m = mem.id;
                  const mt = tasks.filter(t => t.assignee === m && inSelectedWeek(t));
                  const matched = (['ae', 'si', 'pd', 'va', 'pr'] as const).filter(sc => mt.some(t => t.scope === sc));
                  const scArr = matched.length ? matched : (['ae', 'si', 'pd', 'va', 'pr'] as const);

                  return (
                    <div className="member-row" key={m}>
                      <div className="member-info">
                        <span className="av w-8 h-8 font-bold" style={{ background: mem.color, color: '#fff' }}>{memberInitial(m)}</span>
                        <div>
                          <div className="member-name font-bold">{mem.name}</div>
                          <div className="member-role text-[10px] text-muted">{mem.role}</div>
                        </div>
                      </div>
                      
                      <div className="flex gap-4 mr-6 shrink-0 text-center">
                        <div>
                          <div className="text-base font-bold text-main">{mt.length}</div>
                          <div className="text-[9px] uppercase tracking-wide text-muted">tasks</div>
                        </div>
                        <div>
                          <div className="text-base font-bold text-success">{mt.filter(t => t.col === 3).length}</div>
                          <div className="text-[9px] uppercase tracking-wide text-muted">done</div>
                        </div>
                        <div>
                          <div className="text-base font-bold text-warning">{mt.filter(t => t.col === 1).length}</div>
                          <div className="text-[9px] uppercase tracking-wide text-muted">active</div>
                        </div>
                      </div>

                      <div className="wl-bars">
                        {scArr.map(sc => {
                          const st = mt.filter(t => t.scope === sc);
                          const pct = st.length ? Math.round(st.reduce((sum, t) => sum + (t.pct || 0), 0) / st.length) : 0;
                          return (
                            <div className="wl-row" key={sc}>
                              <span className="wl-label font-bold text-muted">{SCOPES[sc].split(' ')[0]}</span>
                              <div className="wl-bar">
                                <div className="wl-fill" style={{ width: `${pct}%`, background: SCOPE_CONFIGS[sc].color }}></div>
                              </div>
                              <span className="wl-pct">{pct}%</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Highlights & Attention */}
              {(() => {
                const overdue = tasks.filter(t => {
                  if (!inSelectedWeek(t)) return false;
                  if (!t.deadline || t.col === 3 || t.col === 4) return false;
                  const d = new Date(t.deadline); d.setHours(0,0,0,0);
                  return d < new Date();
                });
                const hi = tasks.filter(t => inSelectedWeek(t) && t.priority === 'high' && t.col !== 3 && t.col !== 4);

                if (!overdue.length && !hi.length) return null;

                return (
                  <div className="highlight-box">
                    <h4>⚠ Items Needing Attention — Escalated for SUNNY</h4>
                    <ul>
                      {overdue.map(t => (
                        <li key={t.id}><strong>[OVERDUE]</strong> {t.title} — {SCOPES[t.scope]} ({memberName(t.assignee)})</li>
                      ))}
                      {hi.slice(0, 4).map(t => (
                        <li key={t.id}><strong>[HIGH PRIORITY]</strong> {t.title} — {t.pct || 0}% done, deadline: {t.deadline || 'TBD'}</li>
                      ))}
                    </ul>
                  </div>
                );
              })()}

              {/* Next Week Focus */}
              <div className="panel mt-4 no-print">
                <h3 className="mb-2">Next Week Focus & Notes</h3>
                <textarea
                  className="weekly-note"
                  placeholder="Ghi chú định hướng tuần tới, rủi ro cần escalate..."
                  value={weeklyNote}
                  disabled={!isEditMode}
                  onChange={e => setWeeklyNote(e.target.value)}
                />
              </div>

            </div>
          </div>
        )}
      </main>

      {/* --- TASK ADD/EDIT MODAL --- */}
      {isModalOpen && (
        <div className="modal-bg show">
        <div className="modal">
          <h3>
            <span>{editId !== null ? 'Edit Task' : 'Add New Task'}</span>
            <button type="button" className="bg-transparent border-none text-slate-400 hover:text-slate-200 p-1" onClick={closeModal}>
              <X size={16} />
            </button>
          </h3>
          
          <form onSubmit={handleSaveTask}>
            <div className="fg">
              <label>Title *</label>
              <input 
                value={taskForm.title} 
                onChange={e => setTaskForm(prev => ({ ...prev, title: e.target.value }))}
                placeholder="Task title..." 
                required
              />
            </div>
            
            <div className="fg">
              <label>Description</label>
              <textarea
                value={taskForm.desc}
                onChange={e => setTaskForm(prev => ({ ...prev, desc: e.target.value }))}
                placeholder="Brief description..."
              />
            </div>

            <div className="fg">
              <label>Link / Đường dẫn liên kết</label>
              <input
                type="url"
                value={taskForm.link || ''}
                onChange={e => setTaskForm(prev => ({ ...prev, link: e.target.value }))}
                placeholder="https://..."
              />
            </div>

            <div className="fg-row">
              <div className="fg">
                <label>Scope *</label>
                <select 
                  value={taskForm.scope}
                  onChange={e => {
                    const newScope = e.target.value as 'ae' | 'si' | 'pd' | 'va' | 'pr';
                    setTaskForm(prev => ({ 
                      ...prev, 
                      scope: newScope, 
                      subtype: subtypes[newScope][0] || '' 
                    }));
                  }}
                >
                  {(['ae', 'si', 'pd', 'va', 'pr'] as const).map(sc => (
                    <option value={sc} key={sc}>{SCOPES[sc]}</option>
                  ))}
                </select>
              </div>

              <div className="fg">
                <label>Sub-type</label>
                {!isAddingSubtype ? (
                  <>
                    <select 
                      value={taskForm.subtype}
                      onChange={e => setTaskForm(prev => ({ ...prev, subtype: e.target.value }))}
                    >
                      <option value="">— Select —</option>
                      {subtypes[taskForm.scope].map(st => (
                        <option value={st} key={st}>{st}</option>
                      ))}
                    </select>
                    
                    <button 
                      type="button" 
                      className="btn-add-custom-subtype"
                      onClick={() => setIsAddingSubtype(true)}
                    >
                      + Add custom sub-type
                    </button>
                  </>
                ) : (
                  <div className="flex gap-2">
                    <input 
                      placeholder="New subtype name..." 
                      value={newSubtypeInput}
                      onChange={e => setNewSubtypeInput(e.target.value)}
                      className="flex-1"
                    />
                    <button 
                      type="button" 
                      className="btn-p py-1 px-3 border-none font-bold"
                      onClick={handleAddSubtype}
                    >
                      +
                    </button>
                    <button 
                      type="button" 
                      className="btn-cancel-subtype"
                      onClick={() => setIsAddingSubtype(false)}
                    >
                      ✕
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Subtype list management */}
            <div className="fg mt-2">
              <label>Subtypes List Management</label>
              <div className="subtype-list">
                {subtypes[taskForm.scope].map(st => (
                  <div className="subtype-item" key={st}>
                    <span>{st}</span>
                    <button 
                      type="button" 
                      className="text-red-400 border-none bg-transparent hover:text-red-300 font-bold text-xs"
                      onClick={() => handleRemoveSubtype(taskForm.scope, st)}
                    >
                      ✕
                    </button>
                  </div>
                ))}
              </div>
            </div>

            <div className="fg-row">
              <div className="fg">
                <label>Assignee</label>
                <select
                  value={taskForm.assignee}
                  onChange={e => setTaskForm(prev => ({ ...prev, assignee: e.target.value }))}
                >
                  {members.map(mem => (
                    <option value={mem.id} key={mem.id}>{mem.name}</option>
                  ))}
                </select>
              </div>
              <div className="fg">
                <label>Priority</label>
                <select 
                  value={taskForm.priority}
                  onChange={e => setTaskForm(prev => ({ ...prev, priority: e.target.value as 'high' | 'med' | 'low' }))}
                >
                  <option value="high">High</option>
                  <option value="med">Medium</option>
                  <option value="low">Low</option>
                </select>
              </div>
            </div>

            <div className="fg-row">
              <div className="fg">
                <label>Start Date</label>
                <input 
                  type="date" 
                  value={taskForm.start} 
                  onChange={e => setTaskForm(prev => ({ ...prev, start: e.target.value }))}
                />
              </div>
              <div className="fg">
                <label>Deadline</label>
                <input 
                  type="date" 
                  value={taskForm.deadline} 
                  onChange={e => setTaskForm(prev => ({ ...prev, deadline: e.target.value }))}
                />
              </div>
            </div>

            <div className="fg-row">
              <div className="fg">
                <label>% Complete</label>
                <input 
                  type="number" 
                  min="0" 
                  max="100" 
                  value={taskForm.pct} 
                  onChange={e => setTaskForm(prev => ({ ...prev, pct: Math.min(100, Math.max(0, +e.target.value || 0)) }))}
                />
              </div>
              <div className="fg">
                <label>Status</label>
                <select 
                  value={taskForm.col}
                  onChange={e => setTaskForm(prev => ({ ...prev, col: +e.target.value }))}
                >
                  {COLS.map((colName, idx) => (
                    <option value={idx} key={idx}>{colName}</option>
                  ))}
                </select>
              </div>
            </div>

            <div className="fg">
              <label>Progress Note</label>
              <textarea 
                value={taskForm.note} 
                onChange={e => setTaskForm(prev => ({ ...prev, note: e.target.value }))}
                placeholder="Ghi chú tiến độ..." 
                className="min-h-16"
              />
            </div>

            <div className="modal-actions">
              <button type="button" className="hover:text-slate-300" onClick={closeModal}>Cancel</button>
              <button type="submit" className="btn-p">{editId !== null ? 'Save Changes' : 'Add Task'}</button>
            </div>
          </form>
        </div>
        </div>
      )}

      {isMemberModalOpen && (
        <div className="modal-bg show">
          <div className="modal modal-sm">
            <h3>
              <span>{editingMemberId ? `Sửa: ${memberName(editingMemberId)}` : 'Quản lý thành viên'}</span>
              <button type="button" className="bg-transparent border-none text-slate-400 hover:text-slate-200 p-1" onClick={() => { resetMemberForm(); setIsMemberModalOpen(false); }}>
                <X size={16} />
              </button>
            </h3>

            <div className="fg">
              <label>Tên thành viên *</label>
              <input
                value={memberForm.name}
                onChange={e => setMemberForm(prev => ({ ...prev, name: e.target.value }))}
                placeholder="VD: Jupiter"
                onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); handleSubmitMember(); } }}
              />
            </div>

            <div className="fg">
              <label>Vai trò (role)</label>
              <input
                value={memberForm.role}
                onChange={e => setMemberForm(prev => ({ ...prev, role: e.target.value }))}
                placeholder="VD: Creative Lead"
              />
            </div>

            <div className="fg">
              <label>Màu avatar</label>
              <div className="color-swatches">
                {MEMBER_COLORS.map(c => (
                  <button
                    type="button"
                    key={c}
                    className={`color-swatch ${memberForm.color === c ? 'selected' : ''}`}
                    style={{ background: c }}
                    onClick={() => setMemberForm(prev => ({ ...prev, color: c }))}
                  >
                    {memberForm.name.trim().charAt(0).toUpperCase()}
                  </button>
                ))}
              </div>
            </div>

            <div className="modal-actions">
              {editingMemberId && (
                <button type="button" onClick={resetMemberForm}>Hủy</button>
              )}
              <button type="button" className="btn-p" onClick={handleSubmitMember}>
                {editingMemberId ? <><Check size={14} /> Lưu thay đổi</> : <><Plus size={14} /> Thêm thành viên</>}
              </button>
            </div>

            <div className="fg mt-2">
              <label>Danh sách thành viên</label>
              <div className="subtype-list">
                {members.map(mem => (
                  <div className={`subtype-item member-item ${editingMemberId === mem.id ? 'editing' : ''}`} key={mem.id}>
                    <span className="flex items-center gap-2">
                      <span className="av" style={{ background: mem.color, color: '#fff' }}>{memberInitial(mem.id)}</span>
                      <span className="font-semibold">{mem.name}</span>
                      {mem.role && <span className="text-[10px] text-muted">· {mem.role}</span>}
                    </span>
                    <span className="member-item-actions flex items-center gap-2">
                      <button
                        type="button"
                        className="text-slate-400 hover:text-indigo-400 border-none bg-transparent"
                        title="Sửa"
                        onClick={() => startEditMember(mem)}
                      >
                        <Edit2 size={13} />
                      </button>
                      {mem.removable === false ? (
                        <span className="text-[10px] text-muted">Host Lead</span>
                      ) : (
                        <button
                          type="button"
                          className="text-red-400 border-none bg-transparent hover:text-red-300 font-bold text-xs"
                          title="Xóa"
                          onClick={() => handleDeleteMember(mem.id)}
                        >
                          ✕
                        </button>
                      )}
                    </span>
                  </div>
                ))}
              </div>
            </div>

            <div className="modal-actions">
              <button type="button" onClick={() => { resetMemberForm(); setIsMemberModalOpen(false); }}>Đóng</button>
            </div>
          </div>
        </div>
      )}

      {showUnlockModal && (
        <div className="modal-bg show">
          <div className="modal modal-sm">
            <h3>
              <span>Nhập mật khẩu để Chỉnh sửa</span>
              <button type="button" className="bg-transparent border-none text-slate-400 hover:text-slate-200 p-1" onClick={() => setShowUnlockModal(false)}>
                <X size={16} />
              </button>
            </h3>
            <div className="fg">
              <label>Mật khẩu</label>
              <input
                type="password"
                autoFocus
                value={unlockInput}
                onChange={e => { setUnlockInput(e.target.value); setUnlockError(false); }}
                onKeyDown={e => { if (e.key === 'Enter') submitUnlock(); }}
                placeholder="Nhập mật khẩu..."
              />
              {unlockError && <span className="text-[11px] text-danger" style={{ marginTop: 6, display: 'block' }}>Mật khẩu không đúng. Thử lại.</span>}
            </div>
            <div className="modal-actions">
              <button type="button" onClick={() => setShowUnlockModal(false)}>Hủy</button>
              <button type="button" className="btn-p" onClick={submitUnlock}>
                <Unlock size={14} /> Mở khoá
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
