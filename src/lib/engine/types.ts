export const SLA_SECONDS = 2 * 3600;
export const OUTLIER_SECONDS = 24 * 3600;
export const PROPERTY_NAME = "Kimpton Hotel Fontenot";
export const DEFAULT_PM_CYCLE_DAYS = 90;

export type SlaStatus = "on_time" | "breach" | "outlier" | "pending";
export type PerformanceRating =
  | "Excellent"
  | "Good"
  | "SLA Watch"
  | "Critical";

export type WoStatus = "Completed" | "Paused" | "Open" | "In Progress" | string;
export type RequestChannel = "guest" | "staff" | "other";
export type PeriodTag = "In-week" | "Backlog close" | "Aging" | "After window";

export type PmHealth =
  | "completed_week"
  | "on_cycle"
  | "due_week"
  | "overdue"
  | "upcoming"
  | "never";

export interface StaffSourceRow {
  propertyName: string;
  name: string;
  totalWorkOrders: number;
  avgResponseTime: string;
}

export interface WorkOrderSourceRow {
  hotel: string;
  status: string;
  area: string;
  itemType: string;
  issueType: string;
  reportedAt: string | null;
  description: string;
  reportedBy: string;
  requestedBy: string;
  assignedTo: string;
  dueDate: string | null;
  responseTime: string;
  workLoad: string;
  completedAt: string | null;
  completedBy: string;
  resolution: string;
  cost: string;
  notes: string;
}

export interface RoomMasterRow {
  room: string;
  floor: string;
  category: string;
  bedType: string;
  description: string;
}

export interface PmInventoryRow {
  room: string;
  floor: string;
  lastPm: string | null;
  nextDue: string | null;
  status: string;
  assignedTo: string;
  pmType: string;
  cycleDays: number;
  notes: string;
  category: string;
  bedType: string;
  description: string;
}

export interface PmCompletionRow {
  room: string;
  completedAt: string | null;
  completedBy: string;
  hours: string;
  pmType: string;
  notes: string;
  status: string;
}

export interface PmSource {
  master: RoomMasterRow[];
  inventory: PmInventoryRow[];
  weekly: PmCompletionRow[];
  history: PmCompletionRow[];
}

export interface SourceBundle {
  averageResponse: {
    title: string;
    createdFor: string;
    meta: string;
    staff: StaffSourceRow[];
  };
  workOrders: {
    title: string;
    createdFor: string;
    meta: string;
    rows: WorkOrderSourceRow[];
  };
  pm?: PmSource;
}

export interface Period {
  start: string;
  end: string;
  generatedAt: string;
  generatedLabel: string;
  createdFor: string;
  property: string;
  displayRange: string;
}

export interface WorkOrder {
  woNumber: number;
  hotel: string;
  status: string;
  area: string;
  itemType: string;
  issueType: string;
  description: string;
  reportedAt: string | null;
  reportedDate: string;
  reportedBy: string;
  requestedBy: string;
  channel: RequestChannel;
  assignedTo: string;
  dueDate: string | null;
  responseTime: string;
  responseSeconds: number | null;
  workLoad: string;
  workLoadSeconds: number | null;
  actionedAt: string | null;
  actionedDate: string;
  timingNote: string;
  completedAt: string | null;
  completedDate: string;
  completedBy: string;
  resolution: string;
  notes: string;
  cost: string;
  slaStatus: SlaStatus;
  slaLabel: string;
  agingDays: number | null;
  agingBucket: string;
  isHistoric: boolean;
  inPeriod: boolean;
  afterWindow: boolean;
  periodTag: PeriodTag;
}

export interface EngineerKpi {
  name: string;
  totalWorkOrders: number;
  avgResponseTime: string;
  avgResponseSeconds: number | null;
  effectiveAvgExcludingOutliers: string | null;
  effectiveAvgSeconds: number | null;
  slaBreaches: number;
  slaCompliance: number;
  outlierFlags: number;
  rating: PerformanceRating;
  completed: number;
  paused: number;
  open: number;
  inProgress: number;
  totalAssigned: number;
  totalWorkload: string;
  historicClosures: number;
  commentary: string;
  isCore: boolean;
}

export interface IssueSlice {
  itemType: string;
  count: number;
  pct: number;
  members?: Array<{ itemType: string; count: number }>;
}

export interface DailyVolume {
  date: string;
  label: string;
  weekday: string;
  count: number;
}

export interface Recommendation {
  id: number;
  title: string;
  body: string;
  priority: "critical" | "high" | "medium";
}

export interface ValidationCheck {
  id: string;
  name: string;
  pass: boolean;
  expected: string;
  actual: string;
  detail: string;
}

export interface PmRoom {
  room: string;
  floor: string;
  category: string;
  bedType: string;
  description: string;
  lastPm: string | null;
  lastPmLabel: string;
  nextDue: string | null;
  nextDueLabel: string;
  cycleDays: number;
  pmType: string;
  assignedTo: string;
  completedBy: string;
  hours: string;
  notes: string;
  sourceStatus: string;
  health: PmHealth;
  healthLabel: string;
  weekCompleted: boolean;
  ytdCompleted: boolean;
  overdueDays: number | null;
  agingBucket: string;
}

export interface PmEngineer {
  name: string;
  weekCompleted: number;
  ytdCompleted: number;
  hours: string;
  hoursSeconds: number;
}

export interface PmFloor {
  floor: string;
  inventory: number;
  weekCompleted: number;
  ytdCompleted: number;
  overdue: number;
  never: number;
}

export interface PmCategory {
  category: string;
  description: string;
  inventory: number;
  weekCompleted: number;
  ytdCompleted: number;
  overdue: number;
  never: number;
  remaining: number;
}

export interface PmModel {
  loaded: boolean;
  rooms: PmRoom[];
  weekly: PmCompletionRow[];
  history: PmCompletionRow[];
  unmatched: PmCompletionRow[];
  engineers: PmEngineer[];
  floors: PmFloor[];
  categories: PmCategory[];
  daily: DailyVolume[];
  overdue: PmRoom[];
  dueThisWeek: PmRoom[];
  weekRooms: PmRoom[];
  never: PmRoom[];
  ytdRooms: PmRoom[];
  recommendations: Recommendation[];
  brief: string;
  totals: {
    inventory: number;
    weekCompleted: number;
    ytdCompleted: number;
    remaining: number;
    ytdPct: number;
    ytdLabel: string;
    onCycle: number;
    onCyclePct: number;
    overdue: number;
    dueThisWeek: number;
    never: number;
    upcoming: number;
    weeksLeft: number;
    weekTarget: number;
    weekPaceLabel: string;
    rating: PerformanceRating | "—";
    historyCount: number;
    unmatched: number;
  };
  validations: ValidationCheck[];
}

export interface ReportModel {
  period: Period;
  staff: StaffSourceRow[];
  workOrders: WorkOrder[];
  kpis: EngineerKpi[];
  totals: {
    pdfWorkOrders: number;
    detailRecords: number;
    completed: number;
    paused: number;
    open: number;
    inProgress: number;
    slaOnTime: number;
    slaBreaches: number;
    slaBreachesCompleted: number;
    outlierFlags: number;
    slaCompliance: number;
    slaComplianceLabel: string;
    overallRating: PerformanceRating | "—";
    weightedAvgResponse: string;
    unassigned: number;
    guestFacing: number;
    staffFacing: number;
    historicClosures: number;
    afterWindow: number;
  };
  dashboard: {
    byEngineer: Array<{
      name: string;
      completed: number;
      paused: number;
      open: number;
      inProgress: number;
      totalAssigned: number;
    }>;
    issueTypes: IssueSlice[];
    daily: DailyVolume[];
    peakDay: DailyVolume | null;
  };
  breaches: WorkOrder[];
  outliers: WorkOrder[];
  aging: WorkOrder[];
  recommendations: Recommendation[];
  executiveBrief: string;
  validations: ValidationCheck[];
  allValid: boolean;
  pm: PmModel;
}
