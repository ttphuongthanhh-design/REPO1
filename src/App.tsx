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
        body: JSON.stringify({ tasks, subtypes, members }),
      }).catch(() => { /* ignore: localStorage still holds the data */ });
    }, 600);
    return () => clearTimeout(id);
  }, [tasks, subtypes, members]);

  // Member lookup helpers
  const memberById = (id: string) => members.find(m => m.id === id);
  const memberName = (id: string) => memberById(id)?.name ?? id;
  const memberRole = (id: string) => memberById(id)?.role ?? '';
  const memberColor = (id: string) => memberById(id)?.color ?? 'linear-gradient(135deg, #94a3b8, #475569)';
  const memberInitial = (id: string) => (memberById(id)?.name ?? id).trim().charAt(0).toUpperCase() || '?';

  const [activeTab, setActiveTab] = useState<'dash' | 'kanban' | 'timeline' | 'daily' | 'weekly'>('dash');

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
      setTasks(prev => prev.filter(t => t.id !== id));
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

    if (editId !== null) {
      setTasks(prev => prev.map(t => t.id === editId ? { ...t, ...formattedForm } : t));
    } else {
      const nextId = tasks.reduce((max, t) => Math.max(max, t.id), 0) + 1;
      setTasks(prev => [...prev, { id: nextId, ...formattedForm }]);
    }
    setIsModalOpen(false);
  };

  const updateTaskField = (id: number, field: keyof Task, value: any) => {
    if (!isEditMode) return;
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
        return updated;
      }
      return t;
    }));
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

  // Shared task-card renderer used by both the Kanban board and the History Recorded Task box.
  const renderTaskCard = (task: Task, colIdx: number) => (
    <div
      className="task-card animate-fade-in"
      key={task.id}
      draggable={isEditMode}
      onDragStart={() => isEditMode && setDraggedTaskId(task.id)}
      onDragEnd={() => setDraggedTaskId(null)}
    >
      <div className="tc-actions edit-only">
        <button className="flex items-center gap-1 text-[9px]" onClick={() => handleOpenEditModal(task)}>
          <Edit2 size={8} /> Edit
        </button>
        <button className="btn-r flex items-center gap-1 text-[9px] px-2 py-0.5" onClick={() => handleDeleteTask(task.id)}>
          <Trash2 size={8} /> Del
        </button>
      </div>

      <div className="mb-2 flex flex-wrap gap-1">
        <span className={`badge s-${task.scope}`}>{SCOPES[task.scope]}</span>
        {task.subtype && <span className="text-[10px] text-slate-400 self-center">{task.subtype}</span>}
      </div>

      <div className="tc-title font-semibold">{task.title}</div>

      <div className="tc-meta">
        <span className={`badge ${PRIORITY_BADGES[task.priority]}`}>{task.priority}</span>
        <span className="av-tag">
          <span className="av" style={{ background: memberColor(task.assignee), color: '#fff' }}>{memberInitial(task.assignee)}</span>
          {memberName(task.assignee)}
        </span>
      </div>

      {colIdx !== 4 ? (
        <div className="tc-pct">
          <div className="pct-bar">
            <div className="pct-fill" style={{ width: `${task.pct}%` }}></div>
          </div>
          <span className="pct-lbl font-semibold text-slate-400">{task.pct}%</span>
        </div>
      ) : (
        <div className="text-[10px] text-rose-500 font-bold mt-2">✕ Failed / Rejected</div>
      )}

      {colIdx === 3 && (
        <div className="tc-done-date">✓ Done {(task.completedAt || task.deadline) ? `· ${new Date(task.completedAt || task.deadline).toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' })}` : ''}</div>
      )}

      {getDeadlineBadge(task.deadline, colIdx)}

      {task.link && (
        <a
          className="tc-link"
          href={task.link}
          target="_blank"
          rel="noopener noreferrer"
          onClick={e => e.stopPropagation()}
        >
          <LinkIcon size={10} />
          <span className="tc-link-text">{task.link.replace(/^https?:\/\/(www\.)?/, '')}</span>
        </a>
      )}

      {task.note && (
        <div className="text-[10px] text-slate-400 mt-2 bg-slate-950/40 p-2 rounded border-l-2 border-indigo-500 italic">
          "{task.note}"
        </div>
      )}
    </div>
  );

  return (
    <div className={`min-h-screen pb-12 ${isEditMode ? '' : 'read-only'}`}>
      {/* Navigation */}
      <nav className="no-print">
        <div className="nav-logo">IMPACT TEAM</div>
        <div className={`nav-tab ${activeTab === 'dash' ? 'active' : ''}`} onClick={() => setActiveTab('dash')}>Dashboard</div>
        <div className={`nav-tab ${activeTab === 'kanban' ? 'active' : ''}`} onClick={() => setActiveTab('kanban')}>Kanban</div>
        <div className={`nav-tab ${activeTab === 'timeline' ? 'active' : ''}`} onClick={() => setActiveTab('timeline')}>Gantt Timeline</div>
        <div className={`nav-tab ${activeTab === 'daily' ? 'active' : ''}`} onClick={() => setActiveTab('daily')}>Daily Report</div>
        <div className={`nav-tab ${activeTab === 'weekly' ? 'active' : ''}`} onClick={() => setActiveTab('weekly')}>Weekly Report</div>
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
                // Done tasks live in the "History Recorded Task" box below the board.
                if (colIdx === 3) return null;
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

            {/* History Recorded Task — completed (Done) tasks; still counted in Daily & Weekly reports */}
            {(() => {
              const historyTasks = tasks
                .filter(t => {
                  const matchesAssignee = assigneeFilter === 'all' || t.assignee === assigneeFilter;
                  const matchesScope = scopeFilter === 'all' || t.scope === scopeFilter;
                  const matchesSearch = !searchQuery || t.title.toLowerCase().includes(searchQuery.toLowerCase());
                  return t.col === 3 && matchesAssignee && matchesScope && matchesSearch;
                })
                .sort((a, b) => (b.completedAt || b.deadline || '').localeCompare(a.completedAt || a.deadline || ''));

              return (
                <div
                  className="history-box mt-4"
                  onDragOver={e => { e.preventDefault(); e.currentTarget.classList.add('dov'); }}
                  onDragLeave={e => e.currentTarget.classList.remove('dov')}
                  onDrop={e => {
                    e.preventDefault();
                    e.currentTarget.classList.remove('dov');
                    if (draggedTaskId !== null) {
                      updateTaskField(draggedTaskId, 'col', 3);
                      setDraggedTaskId(null);
                    }
                  }}
                >
                  <div className="history-h">
                    <span className="history-title">
                      <span className="col-dot" style={{ background: COL_COLORS[3], color: COL_COLORS[3] }}></span>
                      History Recorded Task
                    </span>
                    <span className="history-sub">Task đã hoàn thành (Done) — vẫn được ghi nhận vào Daily &amp; Weekly Report</span>
                    <span className="col-cnt ml-auto">{historyTasks.length}</span>
                  </div>
                  <div className="history-grid">
                    {historyTasks.length === 0 ? (
                      <div className="history-empty">
                        Chưa có task hoàn thành. Kéo task vào đây hoặc đặt trạng thái "Done" để ghi nhận.
                      </div>
                    ) : (
                      historyTasks.map(task => renderTaskCard(task, 3))
                    )}
                  </div>
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
                        <span className="scope-name border-l-4 pl-2 font-bold" style={{ borderColor: SCOPE_CONFIGS[sc].color }}>{SCOPES[sc]}</span>
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
