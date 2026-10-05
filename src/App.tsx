import {
  useEffect,
  useMemo,
  useState,
  type CSSProperties,
} from "react";
import "./App.css";
import { supabase } from "./lib/supabase";
import * as XLSX from "xlsx";

/* =========================
   الأنواع
========================= */

type VisitStatus =
  | "waiting"
  | "inside"
  | "exited"
  | "cancelled";

type Visit = {
  id: number;
  visitorId: number | null;
  visitNo: string;
  visitorName: string;
  visitorPhone: string;
  visitorEmail: string;
  company: string;
  companyId: number | null;
  employee: string;
  reason: string;
  visitDate: string;
  checkIn: string | null;
  checkOut: string | null;
  status: VisitStatus;
};

type Company = {
  id: number;
  name: string;
  active: boolean;
};

type PreviousVisitor = {
  id: number;
  visitorName: string;
  visitorPhone: string;
  visitorEmail: string;
};

type MailDirection =
  | "incoming"
  | "outgoing";

type MailActionChoice =
  | "signed"
  | "dataEntry"
  | "other";

type MailStatusChoice =
  | "followUp"
  | "waitingSignature"
  | "processed"
  | "preparing"
  | "other";

type MailAttachment = {
  id: number;
  mailId: number;
  fileName: string;
  storagePath: string;
  fileSize: number | null;
  mimeType: string | null;
  createdAt: string;
};

type MailItem = {
  id: number;
  mailNo: string;
  subject: string;
  direction: MailDirection;
  responsible: string;
  action: string;
  status: string;
  mailDate: string;
  mailTime: string;
  attachments: MailAttachment[];
};

type CallStatus = "تحويل مكالمة" | "تم";

type CallItem = {
  id: number;
  callNo: string;
  callerName: string;
  companyName: string;
  phone: string;
  subject: string;
  requestedEmployee: string;
  status: CallStatus;
  createdAt: string;
};

type AppointmentStatus = "قيد الانتظار" | "تم الموعد" | "تأجيل الموعد" | "إلغاء الموعد";
type AppointmentPostponement = "date" | "expired";

type AppointmentItem = {
  id: number;
  appointmentNo: string;
  personName: string;
  dayName: string;
  entryDate: string;
  appointmentDate: string;
  appointmentTime: string;
  hostName: string;
  status: AppointmentStatus;
  postponementType: AppointmentPostponement | null;
  postponedDate: string;
};

/* =========================
   دوال مساعدة
========================= */

function getJordanTime() {
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Amman",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(new Date());
}

function getJordanDateTimeInput() {
  return { date: getJordanDate(), time: getJordanTime() };
}

function getJordanDate() {
  return new Intl.DateTimeFormat(
    "en-CA",
    {
      timeZone: "Asia/Amman",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    },
  ).format(new Date());
}

function getVisitDateKey(value: string | null | undefined) {
  if (!value) return "";
  const raw = String(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return raw.slice(0, 10);

  return new Intl.DateTimeFormat(
    "en-CA",
    {
      timeZone: "Asia/Amman",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    },
  ).format(date);
}


function formatVisitTime(
  value: string | null,
) {
  if (!value) return "—";

  return new Intl.DateTimeFormat(
    "ar-JO",
    {
      timeZone: "Asia/Amman",
      hour: "2-digit",
      minute: "2-digit",
    },
  ).format(new Date(value));
}

function formatCallDateTime(value: string | null | undefined) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";

  return new Intl.DateTimeFormat("ar-JO", {
    timeZone: "Asia/Amman",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatVisitDate(
  value: string,
) {
  if (!value) return "—";

  const [year, month, day] =
    value.split("-");

  return `${day}/${month}/${year}`;
}

function getStatusLabel(
  status: VisitStatus,
) {
  switch (status) {
    case "waiting":
      return "بانتظار الدخول";

    case "inside":
      return "داخل الشركة";

    case "exited":
      return "غادر";

    case "cancelled":
      return "ملغاة";

    default:
      return status;
  }
}

function statusInfo(
  status: VisitStatus,
) {
  switch (status) {
    case "waiting":
      return {
        label: "بانتظار الدخول",
        icon: "↪",
        className: "waiting",
      };

    case "inside":
      return {
        label: "داخل الشركة",
        icon: "●",
        className: "inside",
      };

    case "exited":
      return {
        label: "غادر",
        icon: "↩",
        className: "exited",
      };

    case "cancelled":
      return {
        label: "ملغاة",
        icon: "×",
        className: "cancelled",
      };

    default:
      return {
        label: status,
        icon: "•",
        className: "waiting",
      };
  }
}

function getMailDirectionLabel(
  direction: MailDirection,
) {
  return direction === "incoming"
    ? "وارد"
    : "صادر";
}

function getMailStatusInfo(status: string) {
  switch (status) {
    case "قيد المتابعة":
      return { className: "mail-follow", icon: "↻" };
    case "انتظار التوقيع":
      return { className: "mail-sign", icon: "✎" };
    case "تم المعاملة":
      return { className: "mail-done", icon: "✓" };
    case "قيد الإعداد":
      return { className: "mail-prep", icon: "⚙" };
    default:
      return { className: "mail-other", icon: "•" };
  }
}

function getCallStatusInfo(status: string) {
  return status === "تم"
    ? { className: "call-done", icon: "✓" }
    : { className: "call-transfer", icon: "↗" };
}

function getAppointmentStatusInfo(status: AppointmentStatus) {
  switch (status) {
    case "تم الموعد":
      return { className: "appointment-done", icon: "✓" };
    case "تأجيل الموعد":
      return { className: "appointment-postponed", icon: "↻" };
    case "إلغاء الموعد":
      return { className: "appointment-cancelled", icon: "×" };
    default:
      return { className: "appointment-waiting", icon: "◷" };
  }
}

function getMailDirectionInfo(direction: MailDirection) {
  return direction === "incoming"
    ? { className: "mail-incoming", icon: "↓" }
    : { className: "mail-outgoing", icon: "↑" };
}

function getMailActionInfo(action: string) {
  if (action === "تم توقيع" || action === "تم التوقيع") return { className: "action-signed", icon: "✓" };
  if (action === "قيد الإدخال") return { className: "action-entry", icon: "↓" };
  return { className: "action-other", icon: "•" };
}



function getMailActionChoice(
  value: string,
): {
  choice: MailActionChoice;
  other: string;
} {
  if (value === "تم التوقيع" || value === "تم توقيع") {
    return {
      choice: "signed",
      other: "",
    };
  }

  if (value === "قيد الإدخال") {
    return {
      choice: "dataEntry",
      other: "",
    };
  }

  return {
    choice: "other",
    other: value,
  };
}

function getMailStatusChoice(
  value: string,
): {
  choice: MailStatusChoice;
  other: string;
} {
  if (value === "قيد المتابعة") {
    return {
      choice: "followUp",
      other: "",
    };
  }

  if (value === "انتظار التوقيع") {
    return {
      choice: "waitingSignature",
      other: "",
    };
  }

  if (value === "تم المعاملة") {
    return {
      choice: "processed",
      other: "",
    };
  }

  if (value === "قيد الإعداد") {
    return {
      choice: "preparing",
      other: "",
    };
  }

  return {
    choice: "other",
    other: value,
  };
}

function getHoverColor(hex: string) {
  const normalized =
    hex.replace("#", "");

  if (normalized.length !== 6) {
    return "#1d4ed8";
  }

  const r = Math.max(
    0,
    Math.min(
      255,
      Math.round(
        parseInt(
          normalized.slice(0, 2),
          16,
        ) * 0.84,
      ),
    ),
  );

  const g = Math.max(
    0,
    Math.min(
      255,
      Math.round(
        parseInt(
          normalized.slice(2, 4),
          16,
        ) * 0.84,
      ),
    ),
  );

  const b = Math.max(
    0,
    Math.min(
      255,
      Math.round(
        parseInt(
          normalized.slice(4, 6),
          16,
        ) * 0.84,
      ),
    ),
  );

  return `rgb(${r}, ${g}, ${b})`;
}

/* =========================
   التطبيق
========================= */

function App() {
  /* =========================
     تسجيل الدخول
  ========================= */

  const [isLoggedIn, setIsLoggedIn] =
    useState<boolean>(
      localStorage.getItem("visit_logged_in") === "true",
    );

  const [loginUsername, setLoginUsername] =
    useState("");

  const [loginPassword, setLoginPassword] =
    useState("");

  const [loginError, setLoginError] =
    useState("");

  const [loginLoading, setLoginLoading] =
    useState(false);

  const handleLogin = async () => {
    setLoginError("");

    if (!loginUsername.trim() || !loginPassword) {
      setLoginError(
        "يرجى إدخال اسم المستخدم وكلمة المرور",
      );
      return;
    }

    setLoginLoading(true);

    try {
      const { data, error } =
        await supabase.rpc("check_app_user", {
          p_username: loginUsername.trim(),
          p_password: loginPassword,
        });

      if (error) {
        console.error(error);
        setLoginError(
          "حدث خطأ أثناء تسجيل الدخول",
        );
        return;
      }

      if (!data || data.length === 0) {
        setLoginError(
          "اسم المستخدم أو كلمة المرور غير صحيحة",
        );
        return;
      }

      localStorage.setItem(
        "visit_logged_in",
        "true",
      );

      localStorage.setItem(
        "visit_username",
        data[0].username,
      );

      setIsLoggedIn(true);
      setLoginPassword("");
    } catch (error) {
      console.error(error);
      setLoginError(
        "حدث خطأ أثناء تسجيل الدخول",
      );
    } finally {
      setLoginLoading(false);
    }
  };

  const handleLogout = () => {
    localStorage.removeItem("visit_logged_in");
    localStorage.removeItem("visit_username");

    setIsLoggedIn(false);
    setLoginUsername("");
    setLoginPassword("");
    setLoginError("");
    setCurrentPage("home");
  };

  const [currentPage, setCurrentPage] =
    useState<
      | "home"
      | "visits"
      | "newVisit"
      | "companies"
      | "companyDetails"
      | "visitDetails"
      | "reports"
      | "settings"
      | "mail"
      | "calls"
      | "appointments"
    >("home");

  /* =========================
     اللون
  ========================= */

  const [accentColor, setAccentColor] =
    useState<string>(
      localStorage.getItem(
        "visit_accent_color",
      ) || "#24357d",
    );

  useEffect(() => {
    const root =
      document.documentElement;

    root.style.setProperty(
      "--primary-color",
      accentColor,
    );

    root.style.setProperty(
      "--primary-hover",
      getHoverColor(accentColor),
    );

    root.style.setProperty(
      "--primary-shadow",
      `${accentColor}40`,
    );

    root.style.setProperty(
      "--primary-shadow-soft",
      `${accentColor}33`,
    );

    root.style.setProperty(
      "--primary-focus",
      `${accentColor}14`,
    );

    localStorage.setItem(
      "visit_accent_color",
      accentColor,
    );
  }, [accentColor]);

  /* =========================
     الزيارات
  ========================= */

  const [visits, setVisits] =
    useState<Visit[]>([]);

  const [selectedVisit, setSelectedVisit] =
    useState<Visit | null>(null);

  const [
    loadingVisits,
    setLoadingVisits,
  ] = useState(false);

  const [visitError, setVisitError] =
    useState("");

  const [visitSearch, setVisitSearch] =
    useState("");

  const [
    visitStatusFilter,
    setVisitStatusFilter,
  ] = useState<"all" | VisitStatus>(
    "all",
  );

  /* =========================
     الشركات
  ========================= */

  const [companies, setCompanies] =
    useState<Company[]>([]);

  const [
    loadingCompanies,
    setLoadingCompanies,
  ] = useState(false);

  const [companyError, setCompanyError] =
    useState("");

  const [
    companySearch,
    setCompanySearch,
  ] = useState("");

  const [
    selectedCompany,
    setSelectedCompany,
  ] = useState<string | null>(null);

  const [
    selectedCompanyId,
    setSelectedCompanyId,
  ] = useState<number | null>(null);

  const [
    companyPageSearch,
    setCompanyPageSearch,
  ] = useState("");

  const [
    selectedCompanyPage,
    setSelectedCompanyPage,
  ] = useState<Company | null>(null);

  const [
    companyPageVisits,
    setCompanyPageVisits,
  ] = useState<Visit[]>([]);

  const [
    companyPageVisitors,
    setCompanyPageVisitors,
  ] = useState<
    PreviousVisitor[]
  >([]);

  const [
    loadingCompanyDetails,
    setLoadingCompanyDetails,
  ] = useState(false);

  /* =========================
     الزائر
  ========================= */

  const [
    previousVisitors,
    setPreviousVisitors,
  ] = useState<PreviousVisitor[]>(
    [],
  );

  const [
    loadingVisitors,
    setLoadingVisitors,
  ] = useState(false);

  const [
    visitorError,
    setVisitorError,
  ] = useState("");

  const [
    selectedVisitor,
    setSelectedVisitor,
  ] =
    useState<PreviousVisitor | null>(
      null,
    );

  const [
    showNewVisitor,
    setShowNewVisitor,
  ] = useState(false);

  const [visitorName, setVisitorName] =
    useState("");

  const [visitorPhone, setVisitorPhone] =
    useState("");

  const [visitorEmail, setVisitorEmail] =
    useState("");

  const [
    requestedEmployee,
    setRequestedEmployee,
  ] = useState("");

  const [visitReason, setVisitReason] =
    useState("");

  const [formError, setFormError] =
    useState("");

  const [
    savingVisit,
    setSavingVisit,
  ] = useState(false);

  const [editingVisit, setEditingVisit] = useState(false);
  const [savingVisitEdit, setSavingVisitEdit] = useState(false);
  const [visitEditError, setVisitEditError] = useState("");
  const [editVisitName, setEditVisitName] = useState("");
  const [editVisitPhone, setEditVisitPhone] = useState("");
  const [editVisitEmail, setEditVisitEmail] = useState("");
  const [editVisitCompanyId, setEditVisitCompanyId] = useState<number | null>(null);
  const [editVisitDate, setEditVisitDate] = useState("");
  const [editVisitEmployee, setEditVisitEmployee] = useState("");
  const [editVisitReason, setEditVisitReason] = useState("");

  /* =========================
     نموذج الشركة
  ========================= */

  const [
    showCompanyForm,
    setShowCompanyForm,
  ] = useState(false);

  const [
    companyFormMode,
    setCompanyFormMode,
  ] = useState<"add" | "edit">(
    "add",
  );

  const [
    companyFormFromVisit,
    setCompanyFormFromVisit,
  ] = useState(false);

  const [
    companyFormId,
    setCompanyFormId,
  ] = useState<number | null>(null);

  const [
    companyFormName,
    setCompanyFormName,
  ] = useState("");

  const [
    companyFormActive,
    setCompanyFormActive,
  ] = useState(true);

  const [
    companyFormError,
    setCompanyFormError,
  ] = useState("");

  const [
    savingCompany,
    setSavingCompany,
  ] = useState(false);

  /* =========================
     التقارير
  ========================= */

  const [reportFrom, setReportFrom] =
    useState("");

  const [reportTo, setReportTo] =
    useState("");

  const [
    reportCompany,
    setReportCompany,
  ] = useState("all");

  const [
    reportStatus,
    setReportStatus,
  ] = useState<"all" | VisitStatus>(
    "all",
  );

  const [reportVisitSearch, setReportVisitSearch] = useState("");
  const [companyReportSearch, setCompanyReportSearch] = useState("");
  const [companyReportEmailFilter, setCompanyReportEmailFilter] = useState<"all" | "withEmail" | "withoutEmail">("all");

  // فلاتر التاريخ لتقارير البريد والاتصالات
  const [reportMailFrom, setReportMailFrom] = useState("");
  const [reportMailTo, setReportMailTo] = useState("");
  const [reportCallsFrom, setReportCallsFrom] = useState("");
  const [reportCallsTo, setReportCallsTo] = useState("");

  // فلاتر التاريخ للقوائم الرئيسية
  const [visitDateFrom, setVisitDateFrom] = useState("");
  const [visitDateTo, setVisitDateTo] = useState("");
  const [mailDateFrom, setMailDateFrom] = useState("");
  const [mailDateTo, setMailDateTo] = useState("");
  const [callDateFrom, setCallDateFrom] = useState("");
  const [callDateTo, setCallDateTo] = useState("");

  /* =========================
     البريد
  ========================= */

  const [mailItems, setMailItems] =
    useState<MailItem[]>([]);

  const [
    loadingMail,
    setLoadingMail,
  ] = useState(false);

  const [mailError, setMailError] =
    useState("");

  const [
    showMailForm,
    setShowMailForm,
  ] = useState(false);

  const [
    savingMail,
    setSavingMail,
  ] = useState(false);

  const [
    mailEditingId,
    setMailEditingId,
  ] = useState<number | null>(
    null,
  );

  const [
    mailFormError,
    setMailFormError,
  ] = useState("");

  const [mailSubject, setMailSubject] =
    useState("");

  const [
    mailDirection,
    setMailDirection,
  ] =
    useState<MailDirection>(
      "incoming",
    );

  const [
    mailResponsible,
    setMailResponsible,
  ] = useState("");

  const [
    mailActionChoice,
    setMailActionChoice,
  ] =
    useState<MailActionChoice>(
      "signed",
    );

  const [
    mailActionOther,
    setMailActionOther,
  ] = useState("");

  const [
    mailStatusChoice,
    setMailStatusChoice,
  ] =
    useState<MailStatusChoice>(
      "followUp",
    );

  const [
    mailStatusOther,
    setMailStatusOther,
  ] = useState("");

  const [mailDate, setMailDate] =
    useState(getJordanDate());

  const [mailTime, setMailTime] =
    useState(
      new Date().toLocaleTimeString(
        "en-GB",
        {
          hour: "2-digit",
          minute: "2-digit",
        },
      ),
    );

  const [mailSearch, setMailSearch] =
    useState("");

  const [mailSelectedFiles, setMailSelectedFiles] =
    useState<File[]>([]);

  const [
    mailDirectionFilter,
    setMailDirectionFilter,
  ] = useState<
    "all" | MailDirection
  >("all");

  const [
    mailStatusFilter,
    setMailStatusFilter,
  ] = useState<
    "all" | MailStatusChoice
  >("all");

  /* =========================
     سجل الاتصالات
  ========================= */

  const [callItems, setCallItems] = useState<CallItem[]>([]);
  const [loadingCalls, setLoadingCalls] = useState(false);
  const [callError, setCallError] = useState("");
  const [showCallForm, setShowCallForm] = useState(false);
  const [savingCall, setSavingCall] = useState(false);
  const [callEditingId, setCallEditingId] = useState<number | null>(null);
  const [callFormError, setCallFormError] = useState("");
  const [callerName, setCallerName] = useState("");
  const [callCompanyName, setCallCompanyName] = useState("");
  const [callPhone, setCallPhone] = useState("");
  const [callSubject, setCallSubject] = useState("");
  const [callRequestedEmployee, setCallRequestedEmployee] = useState("");
  const [callStatus, setCallStatus] = useState<CallStatus>("تحويل مكالمة");
  const [callSearch, setCallSearch] = useState("");
  const [callStatusFilter, setCallStatusFilter] = useState<"all" | CallStatus>("all");

  /* =========================
     المواعيد
  ========================= */

  const [appointmentItems, setAppointmentItems] = useState<AppointmentItem[]>([]);
  const [loadingAppointments, setLoadingAppointments] = useState(false);
  const [appointmentError, setAppointmentError] = useState("");
  const [showAppointmentForm, setShowAppointmentForm] = useState(false);
  const [savingAppointment, setSavingAppointment] = useState(false);
  const [appointmentEditingId, setAppointmentEditingId] = useState<number | null>(null);
  const [appointmentFormError, setAppointmentFormError] = useState("");
  const [appointmentPersonName, setAppointmentPersonName] = useState("");
  const [appointmentDayName, setAppointmentDayName] = useState("");
  const [appointmentEntryDate, setAppointmentEntryDate] = useState("");
  const [appointmentDate, setAppointmentDate] = useState("");
  const [appointmentTime, setAppointmentTime] = useState("");
  const [appointmentHostName, setAppointmentHostName] = useState("");
  const [appointmentStatus, setAppointmentStatus] = useState<AppointmentStatus>("قيد الانتظار");
  const [appointmentNewDate, setAppointmentNewDate] = useState("");
  const [appointmentNewTime, setAppointmentNewTime] = useState("");
  const [appointmentSearch, setAppointmentSearch] = useState("");
  const [appointmentStatusFilter, setAppointmentStatusFilter] = useState<"all" | AppointmentStatus>("all");
  const [appointmentDateFrom, setAppointmentDateFrom] = useState("");
  const [appointmentDateTo, setAppointmentDateTo] = useState("");

  const [reportSection, setReportSection] = useState<"visits" | "mail" | "calls" | "appointments" | "companies">("visits");

  /* =========================
     تنسيق الإدخال
  ========================= */

  const formInputStyle: CSSProperties =
    {
      width: "100%",
      height: "44px",
      padding: "0 13px",
      border:
        "1px solid #d1d5db",
      borderRadius: "9px",
      outline: "none",
      fontSize: "14px",
      boxSizing: "border-box",
      background: "#fff",
    };

  /* =========================
     تحميل الشركات
  ========================= */

  const loadCompanies =
    async () => {
      setLoadingCompanies(true);
      setCompanyError("");

      const result =
        await supabase
          .from("companies")
          .select(
            "id, name, active",
          )
          .order("name", {
            ascending: true,
          });

      if (result.error) {
        setCompanyError(result.error.message);
        setLoadingCompanies(false);
        return;
      }

      setCompanies(
        (result.data ?? []).map((row: any) => ({
          id: row.id,
          name: row.name || "",
          active: row.active !== false,
        })),
      );

      setLoadingCompanies(false);
    };

  /* =========================
     تحميل الزيارات
  ========================= */

  const loadVisits = async () => {
    setLoadingVisits(true);
    setVisitError("");

    const companyMap =
      new Map<number, string>();

    companies.forEach(
      (company) => {
        companyMap.set(
          company.id,
          company.name,
        );
      },
    );

    const result = await supabase
      .from("visits")
      .select(
        `
          id,
          visitor_id,
          visit_no,
          visitor_name,
          visitor_phone,
          visitor_email,
          company_id,
          other_company_name,
          requested_employee,
          visit_reason,
          visit_date,
          check_in_at,
          check_out_at,
          status
        `,
      )
      .order("created_at", {
        ascending: false,
      });

    if (result.error) {
      setVisitError(
        result.error.message,
      );
      setLoadingVisits(false);
      return;
    }

    const rows: Visit[] = (
      result.data ?? []
    ).map((row: any) => ({
      id: row.id,

      visitorId: row.visitor_id ?? null,

      visitNo:
        row.visit_no ||
        `V-${String(
          row.id,
        ).padStart(6, "0")}`,

      visitorName:
        row.visitor_name || "",

      visitorPhone:
        row.visitor_phone || "",

      visitorEmail:
        row.visitor_email || "",

      company:
        row.company_id !== null &&
        companyMap.has(row.company_id)
          ? companyMap.get(
              row.company_id,
            ) || ""
          : row.other_company_name ||
            "شخصي",

      companyId:
        row.company_id ?? null,

      employee:
        row.requested_employee ||
        "",

      reason:
        row.visit_reason || "",

      visitDate:
        row.visit_date || "",

      checkIn:
        row.check_in_at || null,

      checkOut:
        row.check_out_at || null,

      status:
        row.status || "waiting",
    }));

    setVisits(rows);
    setLoadingVisits(false);
  };

  /* =========================
     تحميل البريد
  ========================= */

  const loadMail = async () => {
    setLoadingMail(true);
    setMailError("");

    const result = await supabase
      .from("mail")
      .select(
        `
          id,
          mail_no,
          subject,
          direction,
          responsible,
          action,
          status,
          mail_date,
          mail_time,
          created_at
        `,
      )
      .order("created_at", {
        ascending: false,
      });

    if (result.error) {
      setMailError(
        result.error.message,
      );
      setLoadingMail(false);
      return;
    }

    const baseRows = (
      result.data ?? []
    );

    const mailIds = baseRows.map(
      (row: any) => row.id,
    );

    let attachmentsByMail = new Map<
      number,
      MailAttachment[]
    >();

    if (mailIds.length > 0) {
      const attachmentResult =
        await supabase
          .from("mail_attachments")
          .select(
            "id, mail_id, file_name, storage_path, file_size, mime_type, created_at",
          )
          .in("mail_id", mailIds)
          .order("created_at", {
            ascending: false,
          });

      if (!attachmentResult.error) {
        attachmentsByMail =
          new Map<
            number,
            MailAttachment[]
          >();

        for (const row of attachmentResult.data ?? []) {
          const attachment: MailAttachment = {
            id: row.id,
            mailId: row.mail_id,
            fileName: row.file_name || "",
            storagePath: row.storage_path || "",
            fileSize:
              typeof row.file_size === "number"
                ? row.file_size
                : null,
            mimeType: row.mime_type || null,
            createdAt: row.created_at || "",
          };

          const current =
            attachmentsByMail.get(attachment.mailId) || [];
          current.push(attachment);
          attachmentsByMail.set(attachment.mailId, current);
        }
      }
    }

    const rows: MailItem[] =
      baseRows.map((row: any) => ({
        id: row.id,
        mailNo:
          row.mail_no ||
          `M-${String(row.id).padStart(6, "0")}`,
        subject: row.subject || "",
        direction:
          row.direction === "outgoing"
            ? "outgoing"
            : "incoming",
        responsible: row.responsible || "",
        action: row.action || "",
        status: row.status || "",
        mailDate: row.mail_date || "",
        mailTime: row.mail_time
          ? String(row.mail_time).slice(0, 5)
          : "",
        attachments:
          attachmentsByMail.get(row.id) || [],
      }));

    setMailItems(rows);
    setLoadingMail(false);
  };

  /* =========================
     تحميل سجل الاتصالات
  ========================= */

  const loadCalls = async () => {
    setLoadingCalls(true);
    setCallError("");
    const result = await supabase
      .from("calls")
      .select("id, call_no, caller_name, company_name, phone, subject, requested_employee, status, created_at")
      .order("created_at", { ascending: false });

    if (result.error) {
      setCallError(result.error.message);
      setLoadingCalls(false);
      return;
    }

    setCallItems((result.data ?? []).map((row: any) => ({
      id: row.id,
      callNo: row.call_no || `C-${String(row.id).padStart(6, "0")}`,
      callerName: row.caller_name || "",
      companyName: row.company_name || "",
      phone: row.phone || "",
      subject: row.subject || "",
      requestedEmployee: row.requested_employee || "",
      status: row.status === "تم" ? "تم" : "تحويل مكالمة",
      createdAt: row.created_at || "",
    })));
    setLoadingCalls(false);
  };

  /* =========================
     تحميل المواعيد
  ========================= */

  const loadAppointments = async () => {
    setLoadingAppointments(true);
    setAppointmentError("");
    const result = await supabase
      .from("appointments")
      .select("id, appointment_no, person_name, day_name, entry_date, appointment_date, appointment_time, host_name, status, postponement_type, postponed_date, created_at")
      .order("created_at", { ascending: false });

    if (result.error) {
      setAppointmentError(result.error.message);
      setLoadingAppointments(false);
      return;
    }

    setAppointmentItems((result.data ?? []).map((row: any) => ({
      id: row.id,
      appointmentNo: row.appointment_no || `A-${String(row.id).padStart(6, "0")}`,
      personName: row.person_name || "",
      dayName: row.day_name || "",
      entryDate: row.entry_date || "",
      appointmentDate: row.appointment_date || "",
      appointmentTime: row.appointment_time ? String(row.appointment_time).slice(0, 5) : "",
      hostName: row.host_name || "",
      status:
        row.status === "تم"
          ? "تم الموعد"
          : row.status === "تأجيل"
            ? "تأجيل الموعد"
            : row.status === "تم الموعد" ||
                row.status === "تأجيل الموعد" ||
                row.status === "إلغاء الموعد"
              ? row.status
              : "قيد الانتظار",
      postponementType: row.postponement_type === "expired" ? "expired" : row.postponement_type === "date" ? "date" : null,
      postponedDate: row.postponed_date || "",
    })));
    setLoadingAppointments(false);
  };

  /* =========================
     التحميل
  ========================= */

  useEffect(() => {
    void loadCompanies();
  }, []);

  useEffect(() => {
    void loadVisits();
  }, [companies.length]);

  useEffect(() => {
    if (currentPage === "mail" || currentPage === "reports") void loadMail();
    if (currentPage === "calls" || currentPage === "reports") void loadCalls();
    if (currentPage === "appointments" || currentPage === "reports") void loadAppointments();
  }, [currentPage]);

  /* =========================
     الزوار السابقون
  ========================= */

  const loadPreviousVisitors =
    async (
      companyId: number,
    ) => {
      setLoadingVisitors(true);
      setVisitorError("");

      const result =
        await supabase
          .from("visits")
          .select(
            `
              visitor_id,
              visitor_name,
              visitor_phone,
              visitor_email
            `,
          )
          .eq(
            "company_id",
            companyId,
          )
          .not(
            "visitor_id",
            "is",
            null,
          )
          .order(
            "created_at",
            {
              ascending: false,
            },
          );

      if (result.error) {
        setVisitorError(
          result.error.message,
        );
        setLoadingVisitors(false);
        return;
      }

      const seen = new Set<number>();

      const rows: PreviousVisitor[] =
        [];

      (
        result.data ?? []
      ).forEach((row: any) => {
        if (
          !row.visitor_id ||
          seen.has(row.visitor_id)
        ) {
          return;
        }

        seen.add(row.visitor_id);

        rows.push({
          id: row.visitor_id,
          visitorName:
            row.visitor_name ||
            "",
          visitorPhone:
            row.visitor_phone ||
            "",
          visitorEmail:
            row.visitor_email ||
            "",
        });
      });

      setPreviousVisitors(
        rows,
      );

      setLoadingVisitors(false);
    };

  /* =========================
     اختيار الشركة
  ========================= */

  const selectCompany = (
    company: Company,
  ) => {
    setSelectedCompany(
      company.name,
    );

    setSelectedCompanyId(
      company.id,
    );

    setSelectedVisitor(null);
    setShowNewVisitor(false);
    setFormError("");

    void loadPreviousVisitors(
      company.id,
    );
  };

  const selectPersonal = () => {
    setSelectedCompany(
      "شخصي",
    );

    setSelectedCompanyId(null);
    setSelectedVisitor(null);
    setPreviousVisitors([]);
    setShowNewVisitor(true);
    setFormError("");
  };

  const changeCompany = () => {
    setSelectedCompany(null);
    setSelectedCompanyId(null);
    setSelectedVisitor(null);
    setShowNewVisitor(false);
    setPreviousVisitors([]);

    setVisitorName("");
    setVisitorPhone("");
    setVisitorEmail("");
    setRequestedEmployee("");
    setVisitReason("");
    setFormError("");
  };

  const startNewVisitor = () => {
    setSelectedVisitor(null);
    setShowNewVisitor(true);

    setVisitorName("");
    setVisitorPhone("");
    setVisitorEmail("");
    setRequestedEmployee("");
    setVisitReason("");
    setFormError("");
  };

  const selectPreviousVisitor = (
    visitor: PreviousVisitor,
  ) => {
    setSelectedVisitor(
      visitor,
    );

    setVisitorName(
      visitor.visitorName,
    );

    setVisitorPhone(
      visitor.visitorPhone,
    );

    setVisitorEmail(
      visitor.visitorEmail,
    );

    setShowNewVisitor(false);
    setFormError("");
  };

  /* =========================
     الشركة
  ========================= */

  const openAddCompany = (
    fromVisit: boolean,
  ) => {
    setCompanyFormMode("add");
    setCompanyFormId(null);
    setCompanyFormName("");
    setCompanyFormActive(true);
    setCompanyFormError("");
    setCompanyFormFromVisit(
      fromVisit,
    );
    setShowCompanyForm(true);
  };

  const openEditCompany = (
    company: Company,
  ) => {
    setCompanyFormMode("edit");
    setCompanyFormId(company.id);
    setCompanyFormName(
      company.name,
    );
    setCompanyFormActive(
      company.active,
    );
    setCompanyFormError("");
    setCompanyFormFromVisit(false);
    setShowCompanyForm(true);
  };

  const closeCompanyForm = () => {
    if (savingCompany) return;

    setShowCompanyForm(false);
    setCompanyFormError("");
  };

  const saveCompany = async () => {
    const name =
      companyFormName.trim();

    if (!name) {
      setCompanyFormError(
        "أدخل اسم الشركة",
      );
      return;
    }

    setSavingCompany(true);
    setCompanyFormError("");

    if (
      companyFormMode === "add"
    ) {
      const result =
        await supabase
          .from("companies")
          .insert({
            name,
            active:
              companyFormActive,
          })
          .select(
            "id, name, active",
          )
          .single();

      if (result.error) {
        setCompanyFormError(
          result.error.message,
        );
        setSavingCompany(false);
        return;
      }

      const company: Company = {
        id: result.data.id,
        name: result.data.name || "",
        active: result.data.active !== false,
      };

      setCompanies((prev) =>
        [
          ...prev,
          company,
        ].sort((a, b) =>
          a.name.localeCompare(
            b.name,
            "ar",
          ),
        ),
      );

      setShowCompanyForm(
        false,
      );

      if (
        companyFormFromVisit
      ) {
        selectCompany(
          company,
        );
      }
    } else {
      if (
        companyFormId ===
        null
      ) {
        setSavingCompany(false);
        return;
      }

      const result =
        await supabase
          .from("companies")
          .update({
            name,
            active:
              companyFormActive,
          })
          .eq(
            "id",
            companyFormId,
          )
          .select(
            "id, name, active",
          )
          .single();

      if (result.error) {
        setCompanyFormError(
          result.error.message,
        );
        setSavingCompany(false);
        return;
      }

      const updated: Company = {
        id: result.data.id,
        name: result.data.name || "",
        active: result.data.active !== false,
      };

      setCompanies((prev) =>
        prev
          .map((item) =>
            item.id ===
            updated.id
              ? updated
              : item,
          )
          .sort((a, b) =>
            a.name.localeCompare(
              b.name,
              "ar",
            ),
          ),
      );

      setSelectedCompanyPage(
        (prev) =>
          prev &&
          prev.id ===
            updated.id
            ? updated
            : prev,
      );

      setShowCompanyForm(
        false,
      );
    }

    setSavingCompany(false);
  };

  /* =========================
     حفظ زيارة
  ========================= */

  const saveNewVisit = async () => {
    if (
      !visitorName.trim()
    ) {
      setFormError(
        "أدخل اسم الزائر",
      );
      return;
    }

    if (
      !requestedEmployee.trim()
    ) {
      setFormError(
        "أدخل الشخص المطلوب",
      );
      return;
    }

    if (
      !visitReason.trim()
    ) {
      setFormError(
        "أدخل الغرض من الزيارة",
      );
      return;
    }

    setSavingVisit(true);
    setFormError("");

    let visitorId:
      | number
      | null =
      selectedVisitor?.id ??
      null;

    if (!visitorId) {
      const visitorResult =
        await supabase
          .from("visitors")
          .insert({
            name:
              visitorName.trim(),
            phone:
              visitorPhone.trim() ||
              null,
            email:
              visitorEmail.trim() ||
              null,
          })
          .select("id")
          .single();

      if (visitorResult.error) {
        setFormError(
          visitorResult.error.message,
        );

        setSavingVisit(false);
        return;
      }

      visitorId =
        visitorResult.data.id;
    } else {
      await supabase
        .from("visitors")
        .update({
          name:
            visitorName.trim(),
          phone:
            visitorPhone.trim() ||
            null,
          email:
            visitorEmail.trim() ||
            null,
        })
        .eq(
          "id",
          visitorId,
        );
    }

    const visitResult =
      await supabase
        .from("visits")
        .insert({
          visitor_name:
            visitorName.trim(),

          visitor_phone:
            visitorPhone.trim() ||
            null,

          visitor_email:
            visitorEmail.trim() ||
            null,

          company_id:
            selectedCompanyId,

          other_company_name:
            selectedCompanyId ===
            null
              ? "شخصي"
              : null,

          requested_employee:
            requestedEmployee.trim(),

          visit_reason:
            visitReason.trim(),

          visit_date:
            getJordanDate(),

          status: "waiting",

          visitor_id:
            visitorId,
        })
        .select("id")
        .single();

    if (visitResult.error) {
      setFormError(
        visitResult.error.message,
      );

      setSavingVisit(false);
      return;
    }

    const visitId =
      visitResult.data.id;

    const visitNo =
      `V-${String(
        visitId,
      ).padStart(6, "0")}`;

    await supabase
      .from("visits")
      .update({
        visit_no:
          visitNo,
      })
      .eq(
        "id",
        visitId,
      );

    setSavingVisit(false);

    await loadVisits();

    setSelectedCompany(null);
    setSelectedCompanyId(null);
    setSelectedVisitor(null);
    setShowNewVisitor(false);

    setVisitorName("");
    setVisitorPhone("");
    setVisitorEmail("");
    setRequestedEmployee("");
    setVisitReason("");
    setFormError("");

    setCurrentPage("visits");
  };

  /* =========================
     الدخول والخروج
  ========================= */

  const checkIn = async (
    visit: Visit,
  ) => {
    const now =
      new Date().toISOString();

    const result =
      await supabase
        .from("visits")
        .update({
          check_in_at:
            now,
          status: "inside",
        })
        .eq(
          "id",
          visit.id,
        );

    if (result.error) {
      return;
    }

    setVisits((prev) =>
      prev.map((item) =>
        item.id === visit.id
          ? {
              ...item,
              checkIn: now,
              status:
                "inside",
            }
          : item,
      ),
    );
  };

  const checkOut = async (
    visit: Visit,
  ) => {
    const now =
      new Date().toISOString();

    const result =
      await supabase
        .from("visits")
        .update({
          check_out_at:
            now,
          status: "exited",
        })
        .eq(
          "id",
          visit.id,
        );

    if (result.error) {
      return;
    }

    setVisits((prev) =>
      prev.map((item) =>
        item.id === visit.id
          ? {
              ...item,
              checkOut: now,
              status:
                "exited",
            }
          : item,
      ),
    );
  };

  const openVisitDetails = (
    visit: Visit,
  ) => {
    setSelectedVisit(
      visit,
    );

    setCurrentPage(
      "visitDetails",
    );
  };

  const openEditVisit = (visit: Visit) => {
    setSelectedVisit(visit);
    setEditVisitName(visit.visitorName);
    setEditVisitPhone(visit.visitorPhone);
    setEditVisitEmail(visit.visitorEmail);
    setEditVisitCompanyId(visit.companyId);
    setEditVisitDate(visit.visitDate);
    setEditVisitEmployee(visit.employee);
    setEditVisitReason(visit.reason);
    setVisitEditError("");
    setEditingVisit(true);
    setCurrentPage("visitDetails");
  };

  const cancelEditVisit = () => {
    if (savingVisitEdit) return;
    setEditingVisit(false);
    setVisitEditError("");
  };

  const saveVisitEdit = async () => {
    if (!selectedVisit) return;
    if (!editVisitName.trim()) { setVisitEditError("أدخل اسم الزائر"); return; }
    if (!editVisitEmployee.trim()) { setVisitEditError("أدخل الشخص المطلوب"); return; }
    if (!editVisitReason.trim()) { setVisitEditError("أدخل الغرض من الزيارة"); return; }
    if (!editVisitDate) { setVisitEditError("اختر تاريخ الزيارة"); return; }

    setSavingVisitEdit(true);
    setVisitEditError("");
    const visitResult = await supabase
      .from("visits")
      .update({
        visitor_name: editVisitName.trim(),
        visitor_phone: editVisitPhone.trim() || null,
        visitor_email: editVisitEmail.trim() || null,
        company_id: editVisitCompanyId,
        other_company_name: editVisitCompanyId === null ? "شخصي" : null,
        requested_employee: editVisitEmployee.trim(),
        visit_reason: editVisitReason.trim(),
        visit_date: editVisitDate,
      })
      .eq("id", selectedVisit.id);

    if (visitResult.error) {
      setVisitEditError(visitResult.error.message);
      setSavingVisitEdit(false);
      return;
    }

    if (selectedVisit.visitorId) {
      const visitorResult = await supabase
        .from("visitors")
        .update({
          name: editVisitName.trim(),
          phone: editVisitPhone.trim() || null,
          email: editVisitEmail.trim() || null,
        })
        .eq("id", selectedVisit.visitorId);
      if (visitorResult.error) {
        setVisitEditError(visitorResult.error.message);
        setSavingVisitEdit(false);
        return;
      }
    }

    const updated: Visit = {
      ...selectedVisit,
      visitorName: editVisitName.trim(),
      visitorPhone: editVisitPhone.trim(),
      visitorEmail: editVisitEmail.trim(),
      company: editVisitCompanyId === null
        ? "شخصي"
        : companies.find((company) => company.id === editVisitCompanyId)?.name || selectedVisit.company,
      companyId: editVisitCompanyId,
      employee: editVisitEmployee.trim(),
      reason: editVisitReason.trim(),
      visitDate: editVisitDate,
    };
    setSelectedVisit(updated);
    setVisits((prev) => prev.map((item) => item.id === updated.id ? updated : item));
    setCompanyPageVisits((prev) => prev.map((item) => item.id === updated.id ? updated : item));
    setEditingVisit(false);
    setSavingVisitEdit(false);
  };

  /* =========================
     تفاصيل الشركة
  ========================= */

  const openCompanyDetails =
    async (
      company: Company,
    ) => {
      setSelectedCompanyPage(
        company,
      );

      setCurrentPage(
        "companyDetails",
      );

      setLoadingCompanyDetails(
        true,
      );

      const visitResult =
        await supabase
          .from("visits")
          .select(
            `
              id,
              visitor_id,
              visit_no,
              visitor_name,
              visitor_phone,
              visitor_email,
              company_id,
              other_company_name,
              requested_employee,
              visit_reason,
              visit_date,
              check_in_at,
              check_out_at,
              status
            `,
          )
          .eq(
            "company_id",
            company.id,
          )
          .order(
            "created_at",
            {
              ascending:
                false,
            },
          );

      const visitorResult =
        await supabase
          .from("visits")
          .select(
            `
              visitor_id,
              visitor_name,
              visitor_phone,
              visitor_email,
              created_at
            `,
          )
          .eq(
            "company_id",
            company.id,
          )
          .not(
            "visitor_id",
            "is",
            null,
          )
          .order(
            "created_at",
            {
              ascending:
                false,
            },
          );

      if (!visitResult.error) {
        const rows: Visit[] = (
          visitResult.data ??
          []
        ).map((row: any) => ({
          id: row.id,

          visitorId: row.visitor_id ?? null,

          visitNo:
            row.visit_no ||
            `V-${String(
              row.id,
            ).padStart(
              6,
              "0",
            )}`,

          visitorName:
            row.visitor_name ||
            "",

          visitorPhone:
            row.visitor_phone ||
            "",

          visitorEmail:
            row.visitor_email ||
            "",

          company:
            company.name,

          companyId:
            company.id,

          employee:
            row.requested_employee ||
            "",

          reason:
            row.visit_reason ||
            "",

          visitDate:
            row.visit_date ||
            "",

          checkIn:
            row.check_in_at ||
            null,

          checkOut:
            row.check_out_at ||
            null,

          status:
            row.status ||
            "waiting",
        }));

        setCompanyPageVisits(
          rows,
        );
      }

      if (
        !visitorResult.error
      ) {
        const seen = new Set<number>();

        const rows: PreviousVisitor[] =
          [];

        (
          visitorResult.data ??
          []
        ).forEach(
          (row: any) => {
            if (
              !row.visitor_id ||
              seen.has(
                row.visitor_id,
              )
            ) {
              return;
            }

            seen.add(
              row.visitor_id,
            );

            rows.push({
              id: row.visitor_id,

              visitorName:
                row.visitor_name ||
                "",

              visitorPhone:
                row.visitor_phone ||
                "",

              visitorEmail:
                row.visitor_email ||
                "",
            });
          },
        );

        setCompanyPageVisitors(
          rows,
        );
      }

      setLoadingCompanyDetails(
        false,
      );
    };

  const backToCompanies = () => {
    setSelectedCompanyPage(
      null,
    );

    setCompanyPageVisits([]);
    setCompanyPageVisitors([]);

    setCurrentPage(
      "companies",
    );
  };

  const backToVisits = () => {
    setSelectedVisit(null);
    setEditingVisit(false);
    setVisitEditError("");
    setCurrentPage("visits");
  };

  /* =========================
     البريد
  ========================= */

  const resetMailForm = () => {
    setMailEditingId(null);
    setMailSelectedFiles([]);

    setMailSubject("");

    setMailDirection(
      "incoming",
    );

    setMailResponsible("");

    setMailActionChoice(
      "signed",
    );

    setMailActionOther("");

    setMailStatusChoice(
      "followUp",
    );

    setMailStatusOther("");

    setMailDate(
      getJordanDate(),
    );

    setMailTime(
      new Date().toLocaleTimeString(
        "en-GB",
        {
          hour: "2-digit",
          minute: "2-digit",
        },
      ),
    );

    setMailFormError("");
  };

  const openMailForm = () => {
    resetMailForm();
    setShowMailForm(true);
  };

  const openEditMail = (
    mail: MailItem,
  ) => {
    const action =
      getMailActionChoice(
        mail.action,
      );

    const status =
      getMailStatusChoice(
        mail.status,
      );

    setMailEditingId(
      mail.id,
    );

    setMailSubject(
      mail.subject,
    );

    setMailDirection(
      mail.direction,
    );

    setMailResponsible(
      mail.responsible,
    );

    setMailActionChoice(
      action.choice,
    );

    setMailActionOther(
      action.other,
    );

    setMailStatusChoice(
      status.choice,
    );

    setMailStatusOther(
      status.other,
    );

    setMailDate(
      mail.mailDate,
    );

    setMailTime(
      mail.mailTime,
    );

    setMailSelectedFiles([]);
    setMailFormError("");
    setShowMailForm(true);
  };

  const closeMailForm = () => {
    if (savingMail) {
      return;
    }

    setShowMailForm(false);
    resetMailForm();
  };

  const deleteMail = async (mail: MailItem) => {
    if (savingMail) {
      return;
    }

    const confirmed = window.confirm(
      `هل أنت متأكد من حذف البريد ${mail.mailNo}؟\nسيتم حذف المرفقات المرتبطة به أيضًا.`,
    );

    if (!confirmed) {
      return;
    }

    setMailError("");

    const attachmentResult = await supabase
      .from("mail_attachments")
      .select("storage_path")
      .eq("mail_id", mail.id);

    if (attachmentResult.error) {
      setMailError(
        `تعذر تجهيز مرفقات البريد للحذف: ${attachmentResult.error.message}`,
      );
      return;
    }

    const storagePaths = (attachmentResult.data ?? [])
      .map((row: any) => row.storage_path)
      .filter(
        (path: unknown): path is string =>
          typeof path === "string" && path.trim().length > 0,
      );

    if (storagePaths.length > 0) {
      const storageDelete = await supabase.storage
        .from("mail-attachments")
        .remove(storagePaths);

      if (storageDelete.error) {
        setMailError(
          `تعذر حذف مرفقات البريد: ${storageDelete.error.message}`,
        );
        return;
      }
    }

    const deleteResult = await supabase
      .from("mail")
      .delete()
      .eq("id", mail.id);

    if (deleteResult.error) {
      setMailError(
        `تعذر حذف البريد: ${deleteResult.error.message}`,
      );
      return;
    }

    setMailItems((prev) =>
      prev.filter((item) => item.id !== mail.id),
    );
  };

  const deleteMailAttachment = async (
    attachment: MailAttachment,
  ) => {
    if (savingMail) {
      return;
    }

    const confirmed = window.confirm(
      `هل أنت متأكد من حذف المرفق "${attachment.fileName}"؟`,
    );

    if (!confirmed) {
      return;
    }

    setMailFormError("");

    if (attachment.storagePath) {
      const storageDelete = await supabase.storage
        .from("mail-attachments")
        .remove([attachment.storagePath]);

      if (storageDelete.error) {
        setMailFormError(
          `تعذر حذف المرفق من التخزين: ${storageDelete.error.message}`,
        );
        return;
      }
    }

    const deleteResult = await supabase
      .from("mail_attachments")
      .delete()
      .eq("id", attachment.id);

    if (deleteResult.error) {
      setMailFormError(
        `تم حذف الملف من التخزين لكن تعذر حذف سجله: ${deleteResult.error.message}`,
      );
      return;
    }

    setMailItems((current) =>
      current.map((mail) =>
        mail.id === attachment.mailId
          ? {
              ...mail,
              attachments: mail.attachments.filter(
                (item) => item.id !== attachment.id,
              ),
            }
          : mail,
      ),
    );
  };

  const openMailAttachment = async (
    attachment: MailAttachment,
  ) => {
    if (!attachment.storagePath) {
      return;
    }

    const result =
      await supabase.storage
        .from("mail-attachments")
        .createSignedUrl(
          attachment.storagePath,
          600,
        );

    if (result.error || !result.data?.signedUrl) {
      setMailFormError(
        result.error?.message ||
          "تعذر فتح المرفق",
      );
      return;
    }

    window.open(
      result.data.signedUrl,
      "_blank",
      "noopener,noreferrer",
    );
  };

  const uploadMailAttachments = async (
    mailId: number,
  ): Promise<string | null> => {
    if (mailSelectedFiles.length === 0) {
      return null;
    }

    for (const file of mailSelectedFiles) {
      const safeName = file.name
        .replace(/[\\/:*?"<>|]+/g, "_")
        .trim() || "file";

      const storagePath =
        `${mailId}/${Date.now()}-${Math.random().toString(36).slice(2, 9)}-${safeName}`;

      const uploadResult =
        await supabase.storage
          .from("mail-attachments")
          .upload(
            storagePath,
            file,
            {
              upsert: false,
              contentType: file.type || undefined,
            },
          );

      if (uploadResult.error) {
        return uploadResult.error.message;
      }

      const insertResult =
        await supabase
          .from("mail_attachments")
          .insert({
            mail_id: mailId,
            file_name: file.name,
            storage_path: storagePath,
            file_size: file.size,
            mime_type: file.type || null,
          });

      if (insertResult.error) {
        return insertResult.error.message;
      }
    }

    return null;
  };

  const saveMail = async () => {
    if (
      !mailSubject.trim()
    ) {
      setMailFormError(
        "أدخل موضوع البريد",
      );
      return;
    }

    if (
      !mailResponsible.trim()
    ) {
      setMailFormError(
        "أدخل اسم المسؤول",
      );
      return;
    }

    const finalAction =
      mailActionChoice ===
      "signed"
        ? "تم توقيع"
        : mailActionChoice ===
            "dataEntry"
          ? "قيد الإدخال"
          : mailActionOther.trim();

    if (!finalAction) {
      setMailFormError(
        "اكتب الإجراء",
      );
      return;
    }

    const finalStatus =
      mailStatusChoice ===
      "followUp"
        ? "قيد المتابعة"
        : mailStatusChoice ===
            "waitingSignature"
          ? "انتظار التوقيع"
          : mailStatusChoice ===
              "processed"
            ? "تم المعاملة"
            : mailStatusChoice ===
                "preparing"
              ? "قيد الإعداد"
              : mailStatusOther.trim();

    if (!finalStatus) {
      setMailFormError(
        "اكتب الحالة",
      );
      return;
    }

    if (!mailDate) {
      setMailFormError(
        "اختر التاريخ",
      );
      return;
    }

    if (!mailTime) {
      setMailFormError(
        "اختر الوقت",
      );
      return;
    }

    setSavingMail(true);
    setMailFormError("");

    if (
      mailEditingId !== null
    ) {
      const updateResult =
        await supabase
          .from("mail")
          .update({
            subject:
              mailSubject.trim(),
            direction:
              mailDirection,
            responsible:
              mailResponsible.trim(),
            action:
              finalAction,
            status:
              finalStatus,
            mail_date:
              mailDate,
            mail_time:
              mailTime,
          })
          .eq(
            "id",
            mailEditingId,
          );

      if (
        updateResult.error
      ) {
        setMailFormError(
          updateResult.error.message,
        );

        setSavingMail(false);
        return;
      }

      const attachmentError =
        await uploadMailAttachments(
          mailEditingId,
        );

      if (attachmentError) {
        setMailFormError(
          `تم حفظ البريد، لكن تعذر رفع المرفق: ${attachmentError}`,
        );
        await loadMail();
        setSavingMail(false);
        return;
      }

      setMailItems((prev) =>
        prev.map((item) =>
          item.id ===
          mailEditingId
            ? {
                ...item,
                subject:
                  mailSubject.trim(),
                direction:
                  mailDirection,
                responsible:
                  mailResponsible.trim(),
                action:
                  finalAction,
                status:
                  finalStatus,
                mailDate,
                mailTime,
              }
            : item,
        ),
      );

      await loadMail();
      setSavingMail(false);
      setShowMailForm(false);
      resetMailForm();

      return;
    }

    const result =
      await supabase
        .from("mail")
        .insert({
          subject:
            mailSubject.trim(),
          direction:
            mailDirection,
          responsible:
            mailResponsible.trim(),
          action:
            finalAction,
          status:
            finalStatus,
          mail_date:
            mailDate,
          mail_time:
            mailTime,
        })
        .select("id")
        .single();

    if (result.error) {
      setMailFormError(
        result.error.message,
      );

      setSavingMail(false);
      return;
    }

    const id =
      result.data.id;

    const mailNo =
      `M-${String(
        id,
      ).padStart(6, "0")}`;

    const updateResult =
      await supabase
        .from("mail")
        .update({
          mail_no:
            mailNo,
        })
        .eq(
          "id",
          id,
        );

    if (
      updateResult.error
    ) {
      setMailFormError(
        updateResult.error.message,
      );

      setSavingMail(false);
      return;
    }

    const attachmentError =
      await uploadMailAttachments(id);

    if (attachmentError) {
      setMailFormError(
        `تم حفظ البريد، لكن تعذر رفع المرفق: ${attachmentError}`,
      );
      await loadMail();
      setSavingMail(false);
      return;
    }

    const newMail: MailItem = {
      id,
      mailNo,
      subject:
        mailSubject.trim(),
      direction:
        mailDirection,
      responsible:
        mailResponsible.trim(),
      action:
        finalAction,
      status:
        finalStatus,
      mailDate,
      mailTime,
      attachments: [],
    };

    setMailItems(
      (prev) => [
        newMail,
        ...prev,
      ],
    );

    await loadMail();
    setSavingMail(false);
    setShowMailForm(false);

    resetMailForm();
  };

  /* =========================
     سجل الاتصالات - النماذج
  ========================= */

  const resetCallForm = () => {
    setCallEditingId(null);
    setCallerName("");
    setCallCompanyName("");
    setCallPhone("");
    setCallSubject("");
    setCallRequestedEmployee("");
    setCallStatus("تحويل مكالمة");
    setCallFormError("");
  };

  const openCallForm = () => {
    resetCallForm();
    setShowCallForm(true);
  };

  const openEditCall = (call: CallItem) => {
    setCallEditingId(call.id);
    setCallerName(call.callerName);
    setCallCompanyName(call.companyName);
    setCallPhone(call.phone);
    setCallSubject(call.subject);
    setCallRequestedEmployee(call.requestedEmployee);
    setCallStatus(call.status);
    setCallFormError("");
    setShowCallForm(true);
  };

  const closeCallForm = () => {
    if (savingCall) return;
    setShowCallForm(false);
    resetCallForm();
  };

  const saveCall = async () => {
    if (!callerName.trim() || !callCompanyName.trim() || !callPhone.trim() || !callSubject.trim() || !callRequestedEmployee.trim()) {
      setCallFormError("أدخل جميع بيانات الاتصال");
      return;
    }

    setSavingCall(true);
    setCallFormError("");

    const payload = {
      caller_name: callerName.trim(),
      company_name: callCompanyName.trim(),
      phone: callPhone.trim(),
      subject: callSubject.trim(),
      requested_employee: callRequestedEmployee.trim(),
      status: callStatus,
    };

    if (callEditingId !== null) {
      const result = await supabase
        .from("calls")
        .update(payload)
        .eq("id", callEditingId);

      if (result.error) {
        setCallFormError(result.error.message);
        setSavingCall(false);
        return;
      }

      setCallItems((prev) =>
        prev.map((item) =>
          item.id === callEditingId
            ? {
                ...item,
                callerName: callerName.trim(),
                companyName: callCompanyName.trim(),
                phone: callPhone.trim(),
                subject: callSubject.trim(),
                requestedEmployee: callRequestedEmployee.trim(),
                status: callStatus,
              }
            : item,
        ),
      );
    } else {
      const result = await supabase
        .from("calls")
        .insert(payload)
        .select("id, created_at")
        .single();

      if (result.error) {
        setCallFormError(result.error.message);
        setSavingCall(false);
        return;
      }

      const id = result.data.id;
      const callNo = `C-${String(id).padStart(6, "0")}`;

      const numberUpdate = await supabase
        .from("calls")
        .update({ call_no: callNo })
        .eq("id", id);

      if (numberUpdate.error) {
        setCallFormError(numberUpdate.error.message);
        setSavingCall(false);
        return;
      }

      setCallItems((prev) => [
        {
          id,
          callNo,
          callerName: callerName.trim(),
          companyName: callCompanyName.trim(),
          phone: callPhone.trim(),
          subject: callSubject.trim(),
          requestedEmployee: callRequestedEmployee.trim(),
          status: callStatus,
          createdAt: result.data.created_at || "",
        },
        ...prev,
      ]);
    }

    setSavingCall(false);
    setShowCallForm(false);
    resetCallForm();
  };

  /* =========================
     المواعيد - النماذج
  ========================= */

  const resetAppointmentForm = () => {
    setAppointmentEditingId(null);
    setAppointmentPersonName("");
    setAppointmentDayName("");
    const jordanNow = getJordanDateTimeInput();
    setAppointmentEntryDate(jordanNow.date);
    setAppointmentDate(jordanNow.date);
    setAppointmentTime(jordanNow.time);
    setAppointmentHostName("");
    setAppointmentStatus("قيد الانتظار");
    setAppointmentNewDate("");
    setAppointmentNewTime("");
    setAppointmentFormError("");
  };

  const openAppointmentForm = () => {
    resetAppointmentForm();
    setShowAppointmentForm(true);
  };

  const openEditAppointment = (appointment: AppointmentItem) => {
    setAppointmentEditingId(appointment.id);
    setAppointmentPersonName(appointment.personName);
    setAppointmentDayName(appointment.dayName);
    setAppointmentEntryDate(appointment.entryDate);
    setAppointmentDate(appointment.appointmentDate);
    setAppointmentTime(appointment.appointmentTime);
    setAppointmentHostName(appointment.hostName);
    setAppointmentStatus(appointment.status);
    setAppointmentNewDate(appointment.appointmentDate);
    setAppointmentNewTime(appointment.appointmentTime);
    setAppointmentFormError("");
    setShowAppointmentForm(true);
  };

  const closeAppointmentForm = () => {
    if (savingAppointment) return;
    setShowAppointmentForm(false);
    resetAppointmentForm();
  };

  const saveAppointment = async () => {
    if (!appointmentPersonName.trim() || !appointmentDayName.trim() || !appointmentEntryDate || !appointmentDate || !appointmentTime || !appointmentHostName.trim()) {
      setAppointmentFormError("أدخل جميع بيانات الموعد");
      return;
    }

    if (appointmentStatus === "تأجيل الموعد" && (!appointmentNewDate || !appointmentNewTime)) {
      setAppointmentFormError("أدخل التاريخ الجديد والوقت الجديد");
      return;
    }

    setSavingAppointment(true);
    setAppointmentFormError("");

    const finalAppointmentDate =
      appointmentStatus === "تأجيل الموعد"
        ? appointmentNewDate
        : appointmentDate;

    const finalAppointmentTime =
      appointmentStatus === "تأجيل الموعد"
        ? appointmentNewTime
        : appointmentTime;

    const payload = {
      person_name: appointmentPersonName.trim(),
      day_name: appointmentDayName.trim(),
      entry_date: appointmentEntryDate,
      appointment_date: finalAppointmentDate,
      appointment_time: finalAppointmentTime,
      host_name: appointmentHostName.trim(),
      status: appointmentStatus,
    };

    if (appointmentEditingId !== null) {
      const result = await supabase
        .from("appointments")
        .update(payload)
        .eq("id", appointmentEditingId);

      if (result.error) {
        setAppointmentFormError(result.error.message);
        setSavingAppointment(false);
        return;
      }

      setAppointmentItems((prev) =>
        prev.map((item) =>
          item.id === appointmentEditingId
            ? {
                ...item,
                personName: appointmentPersonName.trim(),
                dayName: appointmentDayName.trim(),
                entryDate: appointmentEntryDate,
                appointmentDate: finalAppointmentDate,
                appointmentTime: finalAppointmentTime,
                hostName: appointmentHostName.trim(),
                status: appointmentStatus,
                postponementType: null,
                postponedDate: "",
              }
            : item,
        ),
      );
    } else {
      const result = await supabase
        .from("appointments")
        .insert(payload)
        .select("id")
        .single();

      if (result.error) {
        setAppointmentFormError(result.error.message);
        setSavingAppointment(false);
        return;
      }

      const id = result.data.id;
      const appointmentNo = `A-${String(id).padStart(6, "0")}`;

      const numberUpdate = await supabase
        .from("appointments")
        .update({ appointment_no: appointmentNo })
        .eq("id", id);

      if (numberUpdate.error) {
        setAppointmentFormError(numberUpdate.error.message);
        setSavingAppointment(false);
        return;
      }

      setAppointmentItems((prev) => [
        {
          id,
          appointmentNo,
          personName: appointmentPersonName.trim(),
          dayName: appointmentDayName.trim(),
          entryDate: appointmentEntryDate,
          appointmentDate: finalAppointmentDate,
          appointmentTime: finalAppointmentTime,
          hostName: appointmentHostName.trim(),
          status: appointmentStatus,
          postponementType: null,
          postponedDate: "",
        },
        ...prev,
      ]);
    }

    setSavingAppointment(false);
    setShowAppointmentForm(false);
    resetAppointmentForm();
  };

  /* =========================
     الفلاتر
  ========================= */

  const filteredCompanies =
    useMemo(() => {
      const search =
        companySearch
          .trim()
          .toLowerCase();

      return companies.filter(
        (company) => {
          if (
            !company.active
          ) {
            return false;
          }

          if (!search) {
            return true;
          }

          return company.name
            .toLowerCase()
            .includes(search);
        },
      );
    }, [
      companies,
      companySearch,
    ]);

  const companyVisitCounts =
    useMemo(() => {
      const map =
        new Map<
          number,
          number
        >();

      visits.forEach((visit) => {
        if (
          visit.companyId ===
          null
        ) {
          return;
        }

        map.set(
          visit.companyId,
          (map.get(
            visit.companyId,
          ) ?? 0) + 1,
        );
      });

      return map;
    }, [visits]);

  const filteredCompanyPageCompanies =
    useMemo(() => {
      const search =
        companyPageSearch
          .trim()
          .toLowerCase();

      if (!search) {
        return companies;
      }

      return companies.filter(
        (company) =>
          company.name
            .toLowerCase()
            .includes(search),
      );
    }, [
      companies,
      companyPageSearch,
    ]);

  const filteredVisits =
    useMemo(() => {
      const search =
        visitSearch
          .trim()
          .toLowerCase();

      return visits.filter(
        (visit) => {
          const matchSearch =
            !search ||
            visit.visitNo
              .toLowerCase()
              .includes(search) ||
            visit.visitorName
              .toLowerCase()
              .includes(search) ||
            visit.company
              .toLowerCase()
              .includes(search) ||
            visit.employee
              .toLowerCase()
              .includes(search);

          const matchStatus =
            visitStatusFilter ===
              "all" ||
            visit.status ===
              visitStatusFilter;

          const visitDateKey = getVisitDateKey(visit.visitDate);
          const matchDateFrom =
            !visitDateFrom ||
            visitDateKey >= visitDateFrom;
          const matchDateTo =
            !visitDateTo ||
            visitDateKey <= visitDateTo;

          return (
            matchSearch &&
            matchStatus &&
            matchDateFrom &&
            matchDateTo
          );
        },
      );
    }, [
      visits,
      visitSearch,
      visitStatusFilter,
      visitDateFrom,
      visitDateTo,
    ]);

  const filteredReportVisits =
    useMemo(() => {
      const search = reportVisitSearch.trim().toLowerCase();
      return visits.filter(
        (visit) => {
          const matchSearch = !search || [
            visit.visitNo,
            visit.visitorName,
            visit.visitorPhone,
            visit.visitorEmail,
            visit.company,
            visit.employee,
            visit.reason,
          ].some((value) => value.toLowerCase().includes(search));

          const matchFrom =
            !reportFrom ||
            visit.visitDate >=
              reportFrom;

          const matchTo =
            !reportTo ||
            visit.visitDate <=
              reportTo;

          const matchCompany =
            reportCompany ===
              "all" ||
            String(
              visit.companyId,
            ) ===
              reportCompany;

          const matchStatus =
            reportStatus ===
              "all" ||
            visit.status ===
              reportStatus;

          return (
            matchSearch &&
            matchFrom &&
            matchTo &&
            matchCompany &&
            matchStatus
          );
        },
      );
    }, [
      visits,
      reportFrom,
      reportTo,
      reportCompany,
      reportStatus,
      reportVisitSearch,
    ]);

  const companyReportRows = useMemo(() => {
    const emailsByCompany = new Map<number, Set<string>>();

    for (const visit of visits) {
      if (visit.companyId === null) continue;

      const email = visit.visitorEmail.trim();
      if (!email) continue;

      if (!emailsByCompany.has(visit.companyId)) {
        emailsByCompany.set(visit.companyId, new Set<string>());
      }

      emailsByCompany.get(visit.companyId)!.add(email);
    }

    return companies.map((company) => ({
      id: company.id,
      name: company.name,
      email: Array.from(emailsByCompany.get(company.id) ?? []).join("، "),
    }));
  }, [companies, visits]);

  const filteredCompanyReport = useMemo(() => {
    const search = companyReportSearch.trim().toLowerCase();

    return companyReportRows.filter((company) => {
      const hasEmail = company.email.trim().length > 0;
      const matchSearch =
        !search ||
        company.name.toLowerCase().includes(search) ||
        company.email.toLowerCase().includes(search);

      const matchEmailFilter =
        companyReportEmailFilter === "all" ||
        (companyReportEmailFilter === "withEmail" && hasEmail) ||
        (companyReportEmailFilter === "withoutEmail" && !hasEmail);

      return matchSearch && matchEmailFilter;
    });
  }, [companyReportRows, companyReportSearch, companyReportEmailFilter]);

  const filteredMail =
    useMemo(() => {
      const search =
        mailSearch
          .trim()
          .toLowerCase();

      let statusFilterText:
        | string
        | null = null;

      if (
        mailStatusFilter !==
        "all"
      ) {
        switch (
          mailStatusFilter
        ) {
          case "followUp":
            statusFilterText =
              "قيد المتابعة";
            break;

          case "waitingSignature":
            statusFilterText =
              "انتظار التوقيع";
            break;

          case "processed":
            statusFilterText =
              "تم المعاملة";
            break;

          case "preparing":
            statusFilterText =
              "قيد الإعداد";
            break;
        }
      }

      return mailItems.filter(
        (mail) => {
          const matchSearch =
            !search ||
            mail.mailNo
              .toLowerCase()
              .includes(search) ||
            mail.subject
              .toLowerCase()
              .includes(search) ||
            mail.responsible
              .toLowerCase()
              .includes(search) ||
            mail.action
              .toLowerCase()
              .includes(search) ||
            mail.status
              .toLowerCase()
              .includes(search);

          const matchDirection =
            mailDirectionFilter ===
              "all" ||
            mail.direction ===
              mailDirectionFilter;

          const matchStatus =
            mailStatusFilter ===
              "all" ||
            (mailStatusFilter ===
              "other"
              ? ![
                  "قيد المتابعة",
                  "انتظار التوقيع",
                  "تم المعاملة",
                  "قيد الإعداد",
                ].includes(
                  mail.status,
                )
              : mail.status ===
                statusFilterText);

          const mailDateKey = getVisitDateKey(mail.mailDate);
          const matchDateFrom =
            !mailDateFrom ||
            mailDateKey >= mailDateFrom;
          const matchDateTo =
            !mailDateTo ||
            mailDateKey <= mailDateTo;

          return (
            matchSearch &&
            matchDirection &&
            matchStatus &&
            matchDateFrom &&
            matchDateTo
          );
        },
      );
    }, [
      mailItems,
      mailSearch,
      mailDirectionFilter,
      mailStatusFilter,
      mailDateFrom,
      mailDateTo,
    ]);

  const filteredCalls = useMemo(() => {
    const search = callSearch.trim().toLowerCase();
    return callItems.filter((call) => {
      const matchSearch = !search || [
        call.callNo,
        call.callerName,
        call.companyName,
        call.phone,
        call.subject,
        call.requestedEmployee,
      ].some((value) => value.toLowerCase().includes(search));
      const matchStatus = callStatusFilter === "all" || call.status === callStatusFilter;
      const callDateKey = getVisitDateKey(call.createdAt);
      const matchDateFrom = !callDateFrom || callDateKey >= callDateFrom;
      const matchDateTo = !callDateTo || callDateKey <= callDateTo;
      return matchSearch && matchStatus && matchDateFrom && matchDateTo;
    });
  }, [callItems, callSearch, callStatusFilter, callDateFrom, callDateTo]);

  const filteredAppointments = useMemo(() => {
    const search = appointmentSearch.trim().toLowerCase();
    return appointmentItems.filter((appointment) => {
      const matchSearch = !search || [
        appointment.appointmentNo,
        appointment.personName,
        appointment.dayName,
        appointment.hostName,
      ].some((value) => value.toLowerCase().includes(search));
      const matchStatus = appointmentStatusFilter === "all" || appointment.status === appointmentStatusFilter;
      const matchDateFrom = !appointmentDateFrom || appointment.appointmentDate >= appointmentDateFrom;
      const matchDateTo = !appointmentDateTo || appointment.appointmentDate <= appointmentDateTo;
      return matchSearch && matchStatus && matchDateFrom && matchDateTo;
    });
  }, [appointmentItems, appointmentSearch, appointmentStatusFilter, appointmentDateFrom, appointmentDateTo]);

  /* =========================
     الإحصائيات
  ========================= */

  const waitingCount =
    visits.filter(
      (visit) =>
        visit.status ===
        "waiting",
    ).length;

  const insideCount =
    visits.filter(
      (visit) =>
        visit.status ===
        "inside",
    ).length;

  const exitedCount =
    visits.filter(
      (visit) =>
        visit.status ===
          "exited" &&
        !!visit.checkOut &&
        getVisitDateKey(visit.checkOut) ===
          getJordanDate(),
    ).length;

  const todayVisits =
    visits.filter(
      (visit) =>
        getVisitDateKey(visit.visitDate) ===
        getJordanDate(),
    ).length;

  const exportMailToExcel = () => {
    const data = filteredMail.map((mail) => ({
      "رقم البريد": mail.mailNo,
      "الموضوع": mail.subject,
      "وارد / صادر": getMailDirectionLabel(mail.direction),
      "المسؤول": mail.responsible,
      "الإجراء": mail.action,
      "الحالة": mail.status,
      "التاريخ": mail.mailDate,
      "الوقت": mail.mailTime,
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "البريد");
    XLSX.writeFile(workbook, `تقرير البريد ${getJordanDate()}.xlsx`);
  };

  const exportCallsToExcel = () => {
    const data = filteredCalls.map((call) => ({
      "رقم الاتصال": call.callNo,
      "اسم المتصل": call.callerName,
      "الشركة": call.companyName,
      "الرقم": call.phone,
      "الموضوع": call.subject,
      "الشخص المطلوب": call.requestedEmployee,
      "تاريخ ووقت الاتصال": formatCallDateTime(call.createdAt),
      "الحالة": call.status,
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "الاتصالات");
    XLSX.writeFile(workbook, `تقرير الاتصالات ${getJordanDate()}.xlsx`);
  };

  const exportCompaniesToExcel = () => {
    const data = filteredCompanyReport.map((company) => ({
      "اسم الشركة": company.name,
      "البريد الإلكتروني": company.email || "",
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "الشركات");
    XLSX.writeFile(workbook, `تقرير الشركات ${getJordanDate()}.xlsx`);
  };

  const exportAppointmentsToExcel = () => {
    const data = filteredAppointments.map((appointment) => ({
      "رقم الموعد": appointment.appointmentNo,
      "اسم الشخص": appointment.personName,
      "اليوم": appointment.dayName,
      "التاريخ": appointment.entryDate,
      "وقت الموعد": appointment.appointmentTime,
      "اسم المستضيف": appointment.hostName,
      "تاريخ الموعد": appointment.appointmentDate,
      "الحالة": appointment.status,
      "التاريخ والوقت بعد التأجيل": appointment.status === "تأجيل الموعد" ? `${appointment.appointmentDate} ${appointment.appointmentTime}` : "",
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "المواعيد");
    XLSX.writeFile(workbook, `تقرير المواعيد ${getJordanDate()}.xlsx`);
  };

  /* =========================
     تصدير Excel
  ========================= */

  const exportVisitsToExcel =
    () => {
      const data =
        filteredReportVisits.map(
          (visit) => ({
            "رقم الزيارة":
              visit.visitNo,
            "التاريخ":
              visit.visitDate,
            "الزائر":
              visit.visitorName,
            "الهاتف":
              visit.visitorPhone,
            "البريد الإلكتروني":
              visit.visitorEmail,
            "الشركة":
              visit.company,
            "الشخص المطلوب":
              visit.employee,
            "الغرض":
              visit.reason,
            "وقت الدخول":
              formatVisitTime(
                visit.checkIn,
              ),
            "وقت الخروج":
              formatVisitTime(
                visit.checkOut,
              ),
            "الحالة":
              getStatusLabel(
                visit.status,
              ),
          }),
        );

      const worksheet =
        XLSX.utils.json_to_sheet(
          data,
        );

      const workbook =
        XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "الزيارات",
      );

      XLSX.writeFile(
        workbook,
        `تقرير الزيارات ${getJordanDate()}.xlsx`,
      );
    };

  /* =========================
     الواجهة
  ========================= */

  if (!isLoggedIn) {
    return (
      <div
        className="app-shell login-screen"
        dir="rtl"
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          padding: "24px",
          background:
            "linear-gradient(135deg, #f5f7fb 0%, #eef2f7 100%)",
        }}
      >
        <div
          className="login-card"
          style={{
            width: "100%",
            maxWidth: "430px",
            background: "#ffffff",
            border: "1px solid #e5e7eb",
            borderRadius: "22px",
            padding: "34px",
            boxShadow:
              "0 20px 60px rgba(15, 23, 42, 0.12)",
          }}
        >
          <div
            className="login-logo"
            style={{
              display: "flex",
              justifyContent: "center",
              marginBottom: "22px",
            }}
          >
            <div
              style={{
                width: "92px",
                height: "92px",
                borderRadius: "20px",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                background: "#ffffff",
                border: "1px solid #e5e7eb",
                overflow: "hidden",
              }}
            >
              <img
                src="/shoniz-logo.png"
                alt="شعار شونيز"
                style={{
                  maxWidth: "82px",
                  maxHeight: "82px",
                  objectFit: "contain",
                }}
                onError={(event) => {
                  event.currentTarget.style.display = "none";
                }}
              />
              <span
                style={{
                  fontSize: "20px",
                  fontWeight: 800,
                  color: "#24357d",
                }}
              >
                SHONIZ
              </span>
            </div>
          </div>

          <div
            style={{
              textAlign: "center",
              marginBottom: "28px",
            }}
          >
            <h1
              style={{
                margin: 0,
                fontSize: "25px",
                fontWeight: 800,
                color: "#111827",
              }}
            >
              نظام إدارة الزيارات
            </h1>

            <p
              style={{
                margin: "8px 0 0",
                fontSize: "13px",
                color: "#6b7280",
              }}
            >
              Visitor Management System
            </p>
          </div>

          <div
            className="login-form"
            style={{
              display: "flex",
              flexDirection: "column",
              gap: "16px",
            }}
          >
            <div>
              <label
                htmlFor="login-username"
                style={{
                  display: "block",
                  marginBottom: "8px",
                  fontSize: "13px",
                  fontWeight: 700,
                  color: "#374151",
                }}
              >
                اسم المستخدم
              </label>

              <input
                id="login-username"
                type="text"
                value={loginUsername}
                onChange={(event) =>
                  setLoginUsername(event.target.value)
                }
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    handleLogin();
                  }
                }}
                placeholder="أدخل اسم المستخدم"
                autoComplete="username"
                autoFocus
                disabled={loginLoading}
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  height: "48px",
                  padding: "0 14px",
                  border: "1px solid #d1d5db",
                  borderRadius: "11px",
                  outline: "none",
                  fontSize: "14px",
                  background: "#fff",
                  color: "#111827",
                }}
              />
            </div>

            <div>
              <label
                htmlFor="login-password"
                style={{
                  display: "block",
                  marginBottom: "8px",
                  fontSize: "13px",
                  fontWeight: 700,
                  color: "#374151",
                }}
              >
                كلمة المرور
              </label>

              <input
                id="login-password"
                type="password"
                value={loginPassword}
                onChange={(event) =>
                  setLoginPassword(event.target.value)
                }
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    handleLogin();
                  }
                }}
                placeholder="أدخل كلمة المرور"
                autoComplete="current-password"
                disabled={loginLoading}
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  height: "48px",
                  padding: "0 14px",
                  border: "1px solid #d1d5db",
                  borderRadius: "11px",
                  outline: "none",
                  fontSize: "14px",
                  background: "#fff",
                  color: "#111827",
                }}
              />
            </div>

            {loginError && (
              <div
                style={{
                  padding: "11px 13px",
                  borderRadius: "10px",
                  background: "#fef2f2",
                  border: "1px solid #fecaca",
                  color: "#b91c1c",
                  fontSize: "13px",
                  fontWeight: 600,
                }}
              >
                {loginError}
              </div>
            )}

            <button
              type="button"
              onClick={handleLogin}
              disabled={loginLoading}
              style={{
                width: "100%",
                height: "48px",
                border: "none",
                borderRadius: "11px",
                background: "var(--primary-color, #24357d)",
                color: "#ffffff",
                fontSize: "14px",
                fontWeight: 800,
                cursor: loginLoading
                  ? "not-allowed"
                  : "pointer",
                opacity: loginLoading ? 0.7 : 1,
              }}
            >
              {loginLoading
                ? "جاري تسجيل الدخول..."
                : "تسجيل الدخول"}
            </button>
          </div>

          <div
            style={{
              marginTop: "24px",
              paddingTop: "16px",
              borderTop: "1px solid #f0f0f0",
              textAlign: "center",
              fontSize: "11px",
              color: "#9ca3af",
            }}
          >
            Visitor Management Suite
          </div>
        </div>
      </div>
    );
  }

  return (
    <div
      className="app-shell"
      dir="rtl"
    >
      <aside className="sidebar">
        <div className="sidebar-brand">
          <div className="brand-logo-wrap">
            <img
              src="/shoniz-logo.png"
              alt="شعار شونيز"
              className="brand-logo-image"
              onError={(event) => {
                event.currentTarget.style.display = "none";
              }}
            />
            <span className="brand-logo-fallback">SHONIZ</span>
          </div>

          <div className="brand-copy">
            <strong>
              نظام إدارة الزيارات
            </strong>

            <span>
              Visitor Management System
            </span>
          </div>
        </div>

        <nav className="sidebar-nav">
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "home"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "home",
              )
            }
          >
            <span>⌂</span>
            الرئيسية
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "newVisit"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "newVisit",
              )
            }
          >
            <span>＋</span>
            تسجيل زيارة
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "visits"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "visits",
              )
            }
          >
            <span>▣</span>
            الزيارات
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "mail"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "mail",
              )
            }
          >
            <span>✉</span>
            البريد
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "calls"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "calls",
              )
            }
          >
            <span>☎</span>
            سجل الاتصالات
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "appointments"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "appointments",
              )
            }
          >
            <span>◫</span>
            المواعيد
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "companies"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "companies",
              )
            }
          >
            <span>▤</span>
            الشركات
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "reports"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "reports",
              )
            }
          >
            <span>▥</span>
            التقارير والتصدير
          </button>
          <button
            type="button"
            className={
              "nav-item " +
              (currentPage ===
              "settings"
                ? "active"
                : "")
            }
            onClick={() =>
              setCurrentPage(
                "settings",
              )
            }
          >
            <span>⚙</span>
            الإعدادات
          </button>
        </nav>

        <button
          type="button"
          className="nav-item"
          onClick={handleLogout}
          style={{
            marginTop: "8px",
            color: "#b91c1c",
          }}
        >
          <span>↪</span>
          تسجيل الخروج
        </button>

        <div className="sidebar-developer">
          <span>تطوير النظام</span>
          <strong>Raid M HAYAJNEH</strong>
          <small>Visitor Management Suite</small>
        </div>
      </aside>

      <main className="main-content">
        {/* =========================
            الرئيسية - التصميم الأصلي
        ========================= */}

        {currentPage ===
          "home" && (
          <>
            <header className="topbar">
              <div>
                <h1>
                  الرئيسية
                </h1>

                <p>
                  متابعة وإدارة زيارات اليوم
                </p>
              </div>

              <div className="top-actions">
                <div className="current-date">
                  {new Date().toLocaleDateString(
                    "ar-JO",
                  )}
                </div>

                <button
                  type="button"
                  className="new-visit"
                  onClick={() =>
                    setCurrentPage(
                      "newVisit",
                    )
                  }
                >
                  ＋ تسجيل زيارة جديدة
                </button>
              </div>
            </header>

            <section className="stats">
              <div className="stat-card">
                <div className="stat-content">
                  <span>
                    بانتظار الدخول
                  </span>

                  <strong>
                    {waitingCount}
                  </strong>
                </div>

                <div className="stat-icon waiting-icon">
                  ◷
                </div>
              </div>

              <div className="stat-card">
                <div className="stat-content">
                  <span>
                    داخل الشركة الآن
                  </span>

                  <strong>
                    {insideCount}
                  </strong>
                </div>

                <div className="stat-icon inside-icon">
                  ●
                </div>
              </div>

              <div className="stat-card">
                <div className="stat-content">
                  <span>
                    خرج اليوم
                  </span>

                  <strong>
                    {exitedCount}
                  </strong>
                </div>

                <div className="stat-icon exited-icon">
                  ✓
                </div>
              </div>

              <div className="stat-card">
                <div className="stat-content">
                  <span>
                    إجمالي اليوم
                  </span>

                  <strong>
                    {todayVisits}
                  </strong>
                </div>

                <div className="stat-icon total-icon">
                  ▣
                </div>
              </div>
            </section>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    زيارات اليوم
                  </h2>

                  <p>
                    آخر الزيارات وحالتها الحالية
                  </p>
                </div>

                <button
                  type="button"
                  className="view-all"
                  onClick={() =>
                    setCurrentPage(
                      "visits",
                    )
                  }
                >
                  عرض جميع الزيارات ←
                </button>
              </div>

              <div className="filters">
                <div className="search-box">
                  <span>
                    ⌕
                  </span>

                  <input
                    value={
                      visitSearch
                    }
                    onChange={(event) =>
                      setVisitSearch(
                        event.target
                          .value,
                      )
                    }
                    placeholder="ابحث باسم الزائر أو الشركة..."
                  />
                </div>

                <button
                  type="button"
                  className="filter-button"
                  onClick={() => {
                    const today = getJordanDate();
                    setVisitDateFrom(today);
                    setVisitDateTo(today);
                  }}
                >
                  اليوم ▾
                </button>

                <select
                  value={
                    visitStatusFilter
                  }
                  onChange={(event) =>
                    setVisitStatusFilter(
                      event.target
                        .value as
                        | "all"
                        | VisitStatus,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                    background:
                      "#fff",
                  }}
                >
                  <option value="all">
                    كل الحالات ▾
                  </option>

                  <option value="waiting">
                    بانتظار الدخول
                  </option>

                  <option value="inside">
                    داخل الشركة
                  </option>

                  <option value="exited">
                    خرج
                  </option>

                  <option value="cancelled">
                    ملغاة
                  </option>
                </select>
              </div>

              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>
                        رقم الزيارة
                      </th>

                      <th>
                        الزائر
                      </th>

                      <th>
                        الشركة
                      </th>

                      <th>
                        الموظف المطلوب
                      </th>

                      <th>
                        الغرض من الزيارة
                      </th>

                      <th>
                        الدخول
                      </th>

                      <th>
                        الخروج
                      </th>

                      <th>
                        الحالة
                      </th>

                      <th>
                        الإجراء
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {filteredVisits.map(
                      (visit) => {
                        const status =
                          statusInfo(
                            visit.status,
                          );

                        return (
                          <tr
                            key={
                              visit.id
                            }
                          >
                            <td>
                              <strong>
                                {
                                  visit.visitNo
                                }
                              </strong>
                            </td>

                            <td>
                              {
                                visit.visitorName
                              }
                            </td>

                            <td>
                              {
                                visit.company
                              }
                            </td>

                            <td>
                              {
                                visit.employee
                              }
                            </td>

                            <td>
                              {
                                visit.reason
                              }
                            </td>

                            <td>
                              {visit.checkIn
                                ? formatVisitTime(
                                    visit.checkIn,
                                  )
                                : "—"}
                            </td>

                            <td>
                              {visit.checkOut
                                ? formatVisitTime(
                                    visit.checkOut,
                                  )
                                : "—"}
                            </td>

                            <td>
                              <span
                                className={
                                  "status " +
                                  status.className
                                }
                              >
                                <i>{status.icon}</i>
                                {status.label}
                              </span>
                            </td>

                            <td>
                              {visit.status ===
                                "waiting" && (
                                <button
                                  type="button"
                                  className="action-button check-in"
                                  onClick={() =>
                                    void checkIn(
                                      visit,
                                    )
                                  }
                                >
                                  تسجيل دخول
                                </button>
                              )}

                              {visit.status ===
                                "inside" && (
                                <button
                                  type="button"
                                  className="action-button check-out"
                                  onClick={() =>
                                    void checkOut(
                                      visit,
                                    )
                                  }
                                >
                                  تسجيل خروج
                                </button>
                              )}

                              {visit.status ===
                                "exited" && (
                                <button
                                  type="button"
                                  className="action-button details"
                                  onClick={() =>
                                    openVisitDetails(
                                      visit,
                                    )
                                  }
                                >
                                  عرض التفاصيل
                                </button>
                              )}

                              {visit.status !== "exited" && (
                                <button
                                  type="button"
                                  className="action-button details"
                                  style={{ marginTop: "6px" }}
                                  onClick={() =>
                                    openVisitDetails(
                                      visit,
                                    )
                                  }
                                >
                                  التفاصيل
                                </button>
                              )}
                            </td>
                          </tr>
                        );
                      },
                    )}

                    {filteredVisits.length ===
                      0 && (
                      <tr>
                        <td
                          colSpan={9}
                        >
                          لا توجد زيارات مطابقة
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          </>
        )}

        {/* =========================
            الزيارات
        ========================= */}

        {currentPage ===
          "visits" && (
          <>
            <header className="topbar">
              <div>
                <h1>
                  الزيارات
                </h1>

                <p>
                  إدارة جميع الزيارات
                </p>
              </div>

              <div className="top-actions">
                <button
                  type="button"
                  className="new-visit"
                  onClick={() =>
                    setCurrentPage(
                      "newVisit",
                    )
                  }
                >
                  ＋ تسجيل زيارة
                </button>
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    قائمة الزيارات
                  </h2>

                  <p>
                    البحث ومتابعة حالة الزيارات
                  </p>
                </div>
              </div>

              <div className="filters">
                <div className="search-box">
                  <span>
                    ⌕
                  </span>

                  <input
                    value={
                      visitSearch
                    }
                    onChange={(event) =>
                      setVisitSearch(
                        event.target
                          .value,
                      )
                    }
                    placeholder="ابحث عن رقم الزيارة أو الزائر أو الشركة..."
                  />
                </div>

                <select
                  value={
                    visitStatusFilter
                  }
                  onChange={(event) =>
                    setVisitStatusFilter(
                      event.target
                        .value as
                        | "all"
                        | VisitStatus,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                    background:
                      "#fff",
                  }}
                >
                  <option value="all">
                    كل الحالات
                  </option>

                  <option value="waiting">
                    بانتظار الدخول
                  </option>

                  <option value="inside">
                    داخل الشركة
                  </option>

                  <option value="exited">
                    غادر
                  </option>

                  <option value="cancelled">
                    ملغاة
                  </option>
                </select>

                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>
                  من تاريخ
                  <input type="date" value={visitDateFrom} onChange={(event) => setVisitDateFrom(event.target.value)} aria-label="من تاريخ الزيارات" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} />
                </label>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>
                  إلى تاريخ
                  <input type="date" value={visitDateTo} onChange={(event) => setVisitDateTo(event.target.value)} aria-label="إلى تاريخ الزيارات" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} />
                </label>
              </div>

              {loadingVisits && (
                <div className="empty-state">
                  جاري تحميل الزيارات...
                </div>
              )}

              {visitError && (
                <div className="empty-state">
                  {visitError}
                </div>
              )}

              {!loadingVisits &&
                !visitError && (
                  <div className="table-wrapper">
                    <table>
                      <thead>
                        <tr>
                          <th>
                            رقم الزيارة
                          </th>

                          <th>
                            الزائر
                          </th>

                          <th>
                            البريد الإلكتروني
                          </th>

                          <th>
                            الشركة
                          </th>

                          <th>
                            الشخص المطلوب
                          </th>

                          <th>
                            التاريخ
                          </th>

                          <th>
                            الحالة
                          </th>

                          <th>
                            الإجراء
                          </th>
                        </tr>
                      </thead>

                      <tbody>
                        {filteredVisits.map(
                          (visit) => {
                            const status =
                              statusInfo(
                                visit.status,
                              );

                            return (
                              <tr
                                key={
                                  visit.id
                                }
                              >
                                <td>
                                  <strong>
                                    {
                                      visit.visitNo
                                    }
                                  </strong>
                                </td>

                                <td>
                                  {
                                    visit.visitorName
                                  }
                                </td>

                                <td>
                                  {visit.visitorEmail || "—"}
                                </td>

                                <td>
                                  {
                                    visit.company
                                  }
                                </td>

                                <td>
                                  {
                                    visit.employee
                                  }
                                </td>

                                <td>
                                  {formatVisitDate(
                                    visit.visitDate,
                                  )}
                                </td>

                                <td>
                                  <span
                                    className={
                                      "status " +
                                      status.className
                                    }
                                  >
                                    <i />
                                    {
                                      status.label
                                    }
                                  </span>
                                </td>

                                <td>
                                  <button
                                    type="button"
                                    className="action-button details"
                                    onClick={() =>
                                      openVisitDetails(
                                        visit,
                                      )
                                    }
                                  >
                                    التفاصيل
                                  </button>

                                  <button type="button" className="action-button check-in" style={{marginRight:"6px"}} onClick={() => openEditVisit(visit)}>تعديل</button>

                                  {visit.status ===
                                    "waiting" && (
                                    <button
                                      type="button"
                                      className="action-button check-in"
                                      style={{
                                        marginRight:
                                          "6px",
                                      }}
                                      onClick={() =>
                                        void checkIn(
                                          visit,
                                        )
                                      }
                                    >
                                      دخول
                                    </button>
                                  )}

                                  {visit.status ===
                                    "inside" && (
                                    <button
                                      type="button"
                                      className="action-button check-out"
                                      style={{
                                        marginRight:
                                          "6px",
                                      }}
                                      onClick={() =>
                                        void checkOut(
                                          visit,
                                        )
                                      }
                                    >
                                      خروج
                                    </button>
                                  )}
                                </td>
                              </tr>
                            );
                          },
                        )}

                        {filteredVisits.length ===
                          0 && (
                          <tr>
                            <td
                              colSpan={8}
                            >
                              لا توجد زيارات مطابقة
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
            </section>
          </>
        )}

        {/* =========================
            الشركات
        ========================= */}

        {currentPage ===
          "companies" && (
          <>
            <header className="topbar">
              <div>
                <h1>
                  الشركات
                </h1>

                <p>
                  إدارة الشركات المرتبطة بالزيارات
                </p>
              </div>

              <div className="top-actions">
                <button
                  type="button"
                  className="new-visit"
                  onClick={() =>
                    openAddCompany(
                      false,
                    )
                  }
                >
                  ＋ إضافة شركة
                </button>
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    قائمة الشركات
                  </h2>

                  <p>
                    البحث وعرض حالة الشركات وسجل الزيارات
                  </p>
                </div>

                <div className="current-date">
                  {companies.length} شركة
                </div>
              </div>

              <div className="filters">
                <div className="search-box">
                  <span>
                    ⌕
                  </span>

                  <input
                    value={
                      companyPageSearch
                    }
                    onChange={(event) =>
                      setCompanyPageSearch(
                        event.target
                          .value,
                      )
                    }
                    placeholder="ابحث عن الشركة..."
                  />
                </div>
              </div>

              {loadingCompanies && (
                <div className="empty-state">
                  جاري تحميل الشركات...
                </div>
              )}

              {companyError && (
                <div className="empty-state">
                  {companyError}
                </div>
              )}

              {!loadingCompanies &&
                !companyError && (
                  <div className="table-wrapper">
                    <table>
                      <thead>
                        <tr>
                          <th>
                            الشركة
                          </th>

                          <th>
                            الحالة
                          </th>

                          <th>
                            عدد الزيارات
                          </th>

                          <th>
                            الإجراء
                          </th>
                        </tr>
                      </thead>

                      <tbody>
                        {filteredCompanyPageCompanies.map(
                          (company) => {
                            const visitCount =
                              companyVisitCounts.get(
                                company.id,
                              ) ?? 0;

                            return (
                              <tr
                                key={
                                  company.id
                                }
                              >
                                <td>
                                  <button
                                    type="button"
                                    className="company-name-button"
                                    onClick={() =>
                                      void openCompanyDetails(
                                        company,
                                      )
                                    }
                                  >
                                    {
                                      company.name
                                    }
                                  </button>
                                </td>

                                <td>
                                  <span
                                    className={
                                      "status " +
                                      (company.active
                                        ? "inside"
                                        : "cancelled")
                                    }
                                  >
                                    <i />
                                    {company.active
                                      ? "فعالة"
                                      : "غير فعالة"}
                                  </span>
                                </td>

                                <td>
                                  <strong>
                                    {
                                      visitCount
                                    }
                                  </strong>
                                </td>

                                <td>
                                  <button
                                    type="button"
                                    className="action-button details"
                                    onClick={() =>
                                      void openCompanyDetails(
                                        company,
                                      )
                                    }
                                  >
                                    التفاصيل
                                  </button>

                                  <button
                                    type="button"
                                    className="action-button check-in"
                                    style={{
                                      marginRight:
                                        "6px",
                                    }}
                                    onClick={() =>
                                      openEditCompany(
                                        company,
                                      )
                                    }
                                  >
                                    تعديل
                                  </button>
                                </td>
                              </tr>
                            );
                          },
                        )}

                        {filteredCompanyPageCompanies.length ===
                          0 && (
                          <tr>
                            <td
                              colSpan={4}
                            >
                              لا توجد شركات مطابقة للبحث
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
            </section>
          </>
        )}

        {/* =========================
            تفاصيل الشركة
        ========================= */}

        {currentPage ===
          "companyDetails" &&
          selectedCompanyPage && (
            <>
              <header className="topbar">
                <div>
                  <h1>
                    تفاصيل الشركة
                  </h1>

                  <p>
                    سجل الشركة والزوار والزيارات السابقة
                  </p>
                </div>

                <div className="top-actions">
                  <button
                    type="button"
                    className="view-all"
                    onClick={
                      backToCompanies
                    }
                  >
                    ← العودة للشركات
                  </button>

                  <button
                    type="button"
                    className="new-visit"
                    onClick={() =>
                      openEditCompany(
                        selectedCompanyPage,
                      )
                    }
                  >
                    تعديل الشركة
                  </button>
                </div>
              </header>

              <section className="visits-card">
                <div className="section-header">
                  <div>
                    <h2>
                      {
                        selectedCompanyPage.name
                      }
                    </h2>

                    <p>
                      الشركة رقم #
                      {
                        selectedCompanyPage.id
                      }
                    </p>
                  </div>

                  <span
                    className={
                      "status " +
                      (selectedCompanyPage.active
                        ? "inside"
                        : "cancelled")
                    }
                  >
                    <i />

                    {selectedCompanyPage.active
                      ? "فعالة"
                      : "غير فعالة"}
                  </span>
                </div>

                <div
                  style={{
                    display:
                      "grid",
                    gridTemplateColumns:
                      "repeat(3, minmax(0, 1fr))",
                    gap:
                      "16px",
                    marginTop:
                      "24px",
                  }}
                >
                  <div
                    style={{
                      padding:
                        "18px",
                      border:
                        "1px solid #e5e7eb",
                      borderRadius:
                        "12px",
                      background:
                        "#fff",
                    }}
                  >
                    <div
                      style={{
                        color:
                          "#6b7280",
                        fontSize:
                          "12px",
                        fontWeight:
                          700,
                        marginBottom:
                          "7px",
                      }}
                    >
                      اسم الشركة
                    </div>

                    <strong>
                      {
                        selectedCompanyPage.name
                      }
                    </strong>
                  </div>

                  <div
                    style={{
                      padding:
                        "18px",
                      border:
                        "1px solid #e5e7eb",
                      borderRadius:
                        "12px",
                      background:
                        "#fff",
                    }}
                  >
                    <div
                      style={{
                        color:
                          "#6b7280",
                        fontSize:
                          "12px",
                        fontWeight:
                          700,
                        marginBottom:
                          "7px",
                      }}
                    >
                      إجمالي الزيارات
                    </div>

                    <strong>
                      {
                        companyPageVisits.length
                      }
                    </strong>
                  </div>

                  <div
                    style={{
                      padding:
                        "18px",
                      border:
                        "1px solid #e5e7eb",
                      borderRadius:
                        "12px",
                      background:
                        "#fff",
                    }}
                  >
                    <div
                      style={{
                        color:
                          "#6b7280",
                        fontSize:
                          "12px",
                        fontWeight:
                          700,
                        marginBottom:
                          "7px",
                      }}
                    >
                      عدد الزوار
                    </div>

                    <strong>
                      {
                        companyPageVisitors.length
                      }
                    </strong>
                  </div>
                </div>

                {loadingCompanyDetails ? (
                  <div className="empty-state">
                    جاري تحميل تفاصيل الشركة...
                  </div>
                ) : (
                  <>
                    <div
                      className="section-header"
                      style={{
                        marginTop:
                          "30px",
                      }}
                    >
                      <div>
                        <h2>
                          الزوار السابقون
                        </h2>

                        <p>
                          الأشخاص الذين سبق لهم زيارة هذه الشركة
                        </p>
                      </div>
                    </div>

                    {companyPageVisitors.length >
                    0 ? (
                      <div className="table-wrapper">
                        <table>
                          <thead>
                            <tr>
                              <th>
                                اسم الزائر
                              </th>

                              <th>
                                الهاتف
                              </th>

                              <th>
                                البريد الإلكتروني
                              </th>
                            </tr>
                          </thead>

                          <tbody>
                            {companyPageVisitors.map(
                              (
                                visitor,
                              ) => (
                                <tr
                                  key={
                                    visitor.id
                                  }
                                >
                                  <td>
                                    <strong>
                                      {
                                        visitor.visitorName
                                      }
                                    </strong>
                                  </td>

                                  <td>
                                    {
                                      visitor.visitorPhone ||
                                      "—"
                                    }
                                  </td>

                                  <td>
                                    {
                                      visitor.visitorEmail ||
                                      "—"
                                    }
                                  </td>
                                </tr>
                              ),
                            )}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <div className="empty-state">
                        لا يوجد زوار سابقون لهذه الشركة
                      </div>
                    )}

                    <div
                      className="section-header"
                      style={{
                        marginTop:
                          "30px",
                      }}
                    >
                      <div>
                        <h2>
                          سجل زيارات الشركة
                        </h2>

                        <p>
                          جميع الزيارات المسجلة لهذه الشركة
                        </p>
                      </div>
                    </div>

                    {companyPageVisits.length >
                    0 ? (
                      <div className="table-wrapper">
                        <table>
                          <thead>
                            <tr>
                              <th>
                                رقم الزيارة
                              </th>

                              <th>
                                التاريخ
                              </th>

                              <th>
                                الزائر
                              </th>

                              <th>
                                الشخص المطلوب
                              </th>

                              <th>
                                الغرض
                              </th>

                              <th>
                                الحالة
                              </th>

                              <th>
                                الإجراء
                              </th>
                            </tr>
                          </thead>

                          <tbody>
                            {companyPageVisits.map(
                              (visit) => {
                                const status =
                                  statusInfo(
                                    visit.status,
                                  );

                                return (
                                  <tr
                                    key={
                                      visit.id
                                    }
                                  >
                                    <td>
                                      <strong>
                                        {
                                          visit.visitNo
                                        }
                                      </strong>
                                    </td>

                                    <td>
                                      {formatVisitDate(
                                        visit.visitDate,
                                      )}
                                    </td>

                                    <td>
                                      {
                                        visit.visitorName
                                      }
                                    </td>

                                    <td>
                                      {
                                        visit.employee
                                      }
                                    </td>

                                    <td>
                                      {
                                        visit.reason
                                      }
                                    </td>

                                    <td>
                                      <span
                                        className={
                                          "status " +
                                          status.className
                                        }
                                      >
                                        <i />
                                        {
                                          status.label
                                        }
                                      </span>
                                    </td>

                                    <td>
                                      <button
                                        type="button"
                                        className="action-button details"
                                        onClick={() =>
                                          openVisitDetails(
                                            visit,
                                          )
                                        }
                                      >
                                        عرض التفاصيل
                                      </button>
                                    </td>
                                  </tr>
                                );
                              },
                            )}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <div className="empty-state">
                        لا توجد زيارات مسجلة لهذه الشركة
                      </div>
                    )}
                  </>
                )}

                <div
                  style={{
                    display:
                      "flex",
                    marginTop:
                      "24px",
                  }}
                >
                  <button
                    type="button"
                    className="view-all"
                    onClick={
                      backToCompanies
                    }
                  >
                    ← العودة للشركات
                  </button>
                </div>
              </section>
            </>
          )}

        {/* =========================
            تفاصيل الزيارة
        ========================= */}

        {currentPage ===
          "visitDetails" &&
          selectedVisit && (
            <>
              <header className="topbar">
                <div><h1>تفاصيل الزيارة</h1><p>{editingVisit ? "تعديل معلومات الزيارة" : "جميع بيانات الزيارة"}</p></div>
                <div className="top-actions">
                  {!editingVisit && <button type="button" className="new-visit" onClick={() => openEditVisit(selectedVisit)}>✎ تعديل المعلومات</button>}
                  <button type="button" className="view-all" onClick={backToVisits}>← العودة للزيارات</button>
                </div>
              </header>
              <section className="visits-card">
                <div className="section-header"><div><h2>{selectedVisit.visitNo}</h2><p>{editingVisit ? "تعديل بيانات الزائر والشركة والتاريخ والشخص المطلوب والغرض" : "تفاصيل الزيارة المسجلة"}</p></div><span className={"status " + statusInfo(selectedVisit.status).className}><i />{statusInfo(selectedVisit.status).label}</span></div>

                {!editingVisit ? (
                  <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:"16px",marginTop:"24px"}}>
                    {[
                      ["رقم الزيارة", selectedVisit.visitNo],
                      ["تاريخ الزيارة", formatVisitDate(selectedVisit.visitDate)],
                      ["اسم الزائر", selectedVisit.visitorName],
                      ["رقم الهاتف", selectedVisit.visitorPhone || "—"],
                      ["البريد الإلكتروني", selectedVisit.visitorEmail || "—"],
                      ["الشركة", selectedVisit.company],
                      ["الشخص المطلوب", selectedVisit.employee],
                      ["الغرض من الزيارة", selectedVisit.reason],
                      ["وقت الدخول", selectedVisit.checkIn ? formatVisitTime(selectedVisit.checkIn) : "لم يتم الدخول"],
                      ["وقت الخروج", selectedVisit.checkOut ? formatVisitTime(selectedVisit.checkOut) : "لم يتم الخروج"],
                    ].map(([label,value]) => <div key={label} style={{padding:"18px",border:"1px solid #e5e7eb",borderRadius:"12px",background:"#fff"}}><div style={{color:"#6b7280",fontSize:"12px",fontWeight:700,marginBottom:"7px"}}>{label}</div><strong style={{wordBreak:"break-word"}}>{value}</strong></div>)}
                  </div>
                ) : (
                  <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:"16px",marginTop:"24px"}}>
                    <div><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>اسم الزائر *</label><input value={editVisitName} onChange={(event)=>setEditVisitName(event.target.value)} style={formInputStyle} /></div>
                    <div><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>رقم الهاتف</label><input value={editVisitPhone} onChange={(event)=>setEditVisitPhone(event.target.value)} style={formInputStyle} /></div>
                    <div><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>البريد الإلكتروني</label><input type="email" value={editVisitEmail} onChange={(event)=>setEditVisitEmail(event.target.value)} style={formInputStyle} /></div>
                    <div><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>تاريخ الزيارة *</label><input type="date" value={editVisitDate} onChange={(event)=>setEditVisitDate(event.target.value)} style={formInputStyle} /></div>
                    <div><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>الشركة *</label><select value={editVisitCompanyId === null ? "personal" : String(editVisitCompanyId)} onChange={(event)=>setEditVisitCompanyId(event.target.value === "personal" ? null : Number(event.target.value))} style={formInputStyle}><option value="personal">شخصي</option>{companies.filter((company) => company.active).map((company) => <option key={company.id} value={company.id}>{company.name}</option>)}</select></div>
                    <div><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>الشخص المطلوب *</label><input value={editVisitEmployee} onChange={(event)=>setEditVisitEmployee(event.target.value)} style={formInputStyle} /></div>
                    <div style={{gridColumn:"1 / -1"}}><label style={{display:"block",marginBottom:"8px",fontSize:"13px",fontWeight:700,color:"#374151"}}>الغرض من الزيارة *</label><input value={editVisitReason} onChange={(event)=>setEditVisitReason(event.target.value)} style={formInputStyle} /></div>
                  </div>
                )}

                {visitEditError && editingVisit && <div style={{marginTop:"16px",padding:"11px 14px",borderRadius:"9px",background:"#fef2f2",border:"1px solid #fecaca",color:"#b91c1c",fontSize:"13px",fontWeight:600}}>{visitEditError}</div>}
                <div style={{display:"flex",gap:"10px",marginTop:"24px"}}>
                  {editingVisit ? <><button type="button" className="view-all" disabled={savingVisitEdit} onClick={cancelEditVisit}>إلغاء</button><button type="button" className="new-visit" disabled={savingVisitEdit} onClick={()=>void saveVisitEdit()}>{savingVisitEdit ? "جاري الحفظ..." : "حفظ التعديل"}</button></> : <button type="button" className="view-all" onClick={backToVisits}>← العودة للزيارات</button>}
                </div>
              </section>
            </>
          )}

        {/* =========================
            سجل الاتصالات
        ========================= */}

        {currentPage === "calls" && (
          <>
            <header className="topbar">
              <div>
                <h1>سجل الاتصالات</h1>
                <p>تسجيل ومتابعة الاتصالات</p>
              </div>
              <div className="top-actions">
                <div className="current-date">{callItems.length} اتصال</div>
                <button type="button" className="new-visit" onClick={openCallForm}>＋ إضافة اتصال</button>
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>قائمة الاتصالات</h2>
                  <p>جميع الاتصالات المسجلة</p>
                </div>
              </div>

              <div className="filters">
                <div className="search-box">
                  <span>⌕</span>
                  <input value={callSearch} onChange={(event) => setCallSearch(event.target.value)} placeholder="ابحث باسم المتصل أو الشركة أو الموضوع..." />
                </div>
                <select value={callStatusFilter} onChange={(event) => setCallStatusFilter(event.target.value as "all" | CallStatus)} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}>
                  <option value="all">كل الحالات</option>
                  <option value="تحويل مكالمة">تحويل مكالمة</option>
                  <option value="تم">تم</option>
                </select>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>من تاريخ <input type="date" value={callDateFrom} onChange={(event) => setCallDateFrom(event.target.value)} aria-label="من تاريخ الاتصالات" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>إلى تاريخ <input type="date" value={callDateTo} onChange={(event) => setCallDateTo(event.target.value)} aria-label="إلى تاريخ الاتصالات" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
              </div>

              {loadingCalls && <div className="empty-state">جاري تحميل الاتصالات...</div>}
              {callError && <div className="empty-state">{callError}</div>}

              {!loadingCalls && !callError && (
                <div className="table-wrapper">
                  <table>
                    <thead>
                      <tr>
                        <th>رقم الاتصال</th>
                        <th>اسم المتصل</th>
                        <th>الشركة</th>
                        <th>الرقم</th>
                        <th>الموضوع</th>
                        <th>الشخص المطلوب</th>
                        <th>تاريخ ووقت الاتصال</th>
                        <th>الحالة</th>
                        <th>الإجراء</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredCalls.map((call) => (
                        <tr key={call.id}>
                          <td><strong>{call.callNo}</strong></td>
                          <td>{call.callerName}</td>
                          <td>{call.companyName}</td>
                          <td>{call.phone}</td>
                          <td>{call.subject}</td>
                          <td>{call.requestedEmployee}</td>
                          <td>{formatCallDateTime(call.createdAt)}</td>
                          <td>{(() => { const info = getCallStatusInfo(call.status); return <span className={"status " + info.className}><i>{info.icon}</i>{call.status}</span>; })()}</td>
                          <td><button type="button" className="action-button details" onClick={() => openEditCall(call)}>تعديل</button></td>
                        </tr>
                      ))}
                      {filteredCalls.length === 0 && <tr><td colSpan={9}>لا توجد اتصالات مطابقة</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </>
        )}

        {/* =========================
            المواعيد
        ========================= */}

        {currentPage === "appointments" && (
          <>
            <header className="topbar">
              <div>
                <h1>المواعيد</h1>
                <p>تسجيل ومتابعة المواعيد</p>
              </div>
              <div className="top-actions">
                <div className="current-date">{appointmentItems.length} موعد</div>
                <button type="button" className="new-visit" onClick={openAppointmentForm}>＋ إضافة موعد</button>
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>قائمة المواعيد</h2>
                  <p>جميع المواعيد المسجلة</p>
                </div>
              </div>

              <div className="filters">
                <div className="search-box">
                  <span>⌕</span>
                  <input value={appointmentSearch} onChange={(event) => setAppointmentSearch(event.target.value)} placeholder="ابحث باسم الشخص أو المستضيف..." />
                </div>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>من تاريخ <input type="date" value={appointmentDateFrom} onChange={(event) => setAppointmentDateFrom(event.target.value)} aria-label="من تاريخ" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>إلى تاريخ <input type="date" value={appointmentDateTo} onChange={(event) => setAppointmentDateTo(event.target.value)} aria-label="إلى تاريخ" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                <select value={appointmentStatusFilter} onChange={(event) => setAppointmentStatusFilter(event.target.value as "all" | AppointmentStatus)} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}>
                  <option value="all">كل الحالات</option>
                  <option value="قيد الانتظار">قيد الانتظار</option>
                  <option value="تم الموعد">تم الموعد</option>
                  <option value="تأجيل الموعد">تأجيل الموعد</option>
                  <option value="إلغاء الموعد">إلغاء الموعد</option>
                </select>
              </div>

              {loadingAppointments && <div className="empty-state">جاري تحميل المواعيد...</div>}
              {appointmentError && <div className="empty-state">{appointmentError}</div>}

              {!loadingAppointments && !appointmentError && (
                <div className="table-wrapper">
                  <table>
                    <thead>
                      <tr>
                        <th>رقم الموعد</th>
                        <th>اسم الشخص</th>
                        <th>اليوم</th>
                        <th>التاريخ</th>
                        <th>وقت الموعد</th>
                        <th>اسم المستضيف</th>
                        <th>تاريخ الموعد</th>
                        <th>الحالة</th>
                        <th>التأجيل</th>
                        <th>الإجراء</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredAppointments.map((appointment) => (
                        <tr key={appointment.id}>
                          <td><strong>{appointment.appointmentNo}</strong></td>
                          <td>{appointment.personName}</td>
                          <td>{appointment.dayName}</td>
                          <td>{formatVisitDate(appointment.entryDate)}</td>
                          <td>{appointment.appointmentTime}</td>
                          <td>{appointment.hostName}</td>
                          <td>{formatVisitDate(appointment.appointmentDate)}</td>
                          <td>{(() => { const info = getAppointmentStatusInfo(appointment.status); return <span className={"status " + info.className}><i>{info.icon}</i>{appointment.status}</span>; })()}</td>
                          <td>{appointment.status === "تأجيل الموعد" ? `${formatVisitDate(appointment.appointmentDate)} ${appointment.appointmentTime}` : "—"}</td>
                          <td><button type="button" className="action-button details" onClick={() => openEditAppointment(appointment)}>تعديل</button></td>
                        </tr>
                      ))}
                      {filteredAppointments.length === 0 && <tr><td colSpan={10}>لا توجد مواعيد مطابقة</td></tr>}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </>
        )}

        {/* =========================
            التقارير
        ========================= */}

        {currentPage === "reports" && (
          <>
            <header className="topbar">
              <div>
                <h1>التقارير والتصدير</h1>
                <p>استخراج بيانات الزيارات حسب الفلاتر</p>
              </div>
              <div className="top-actions">
                <select value={reportSection} onChange={(event) => setReportSection(event.target.value as "visits" | "mail" | "calls" | "appointments" | "companies")} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}>
                  <option value="visits">تقرير الزيارات</option>
                  <option value="mail">تقرير البريد</option>
                  <option value="calls">تقرير الاتصالات</option>
                  <option value="appointments">تقرير المواعيد</option>
                  <option value="companies">تقرير الشركات</option>
                </select>
                <button type="button" className="new-visit" onClick={() => { if (reportSection === "mail") exportMailToExcel(); else if (reportSection === "calls") exportCallsToExcel(); else if (reportSection === "appointments") exportAppointmentsToExcel(); else if (reportSection === "companies") exportCompaniesToExcel(); else exportVisitsToExcel(); }}>تصدير Excel</button>
              </div>
            </header>

            <div style={{display:"flex",gap:"8px",flexWrap:"wrap",marginBottom:"16px"}}>
              {([
                ["visits", "الزيارات"],
                ["mail", "البريد"],
                ["calls", "الاتصالات"],
                ["appointments", "المواعيد"],
                ["companies", "الشركات"],
              ] as const).map(([value, label]) => (
                <button key={value} type="button" className={value === reportSection ? "new-visit" : "filter-button"} onClick={() => setReportSection(value)}>{label}</button>
              ))}
            </div>

            {reportSection === "visits" && (
            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    فلاتر التقرير
                  </h2>

                  <p>
                    حدد البيانات التي تريد تصديرها
                  </p>
                </div>
              </div>

              <div className="filters">
                <div className="search-box"><span>⌕</span><input value={reportVisitSearch} onChange={(event) => setReportVisitSearch(event.target.value)} placeholder="ابحث برقم الزيارة أو الزائر أو الشركة أو الموظف..." /></div>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>
                  من تاريخ
                  <input
                  type="date"
                  value={
                    reportFrom
                  }
                  onChange={(event) =>
                    setReportFrom(
                      event.target
                        .value,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                  }}
                />
                </label>

                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>
                  إلى تاريخ
                  <input
                  type="date"
                  value={
                    reportTo
                  }
                  onChange={(event) =>
                    setReportTo(
                      event.target
                        .value,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                  }}
                />
                </label>

                <select
                  value={
                    reportCompany
                  }
                  onChange={(event) =>
                    setReportCompany(
                      event.target
                        .value,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                  }}
                >
                  <option value="all">
                    كل الشركات
                  </option>

                  {companies.map(
                    (company) => (
                      <option
                        key={
                          company.id
                        }
                        value={
                          company.id
                        }
                      >
                        {
                          company.name
                        }
                      </option>
                    ),
                  )}
                </select>

                <select
                  value={
                    reportStatus
                  }
                  onChange={(event) =>
                    setReportStatus(
                      event.target
                        .value as
                        | "all"
                        | VisitStatus,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                  }}
                >
                  <option value="all">
                    كل الحالات
                  </option>

                  <option value="waiting">
                    بانتظار الدخول
                  </option>

                  <option value="inside">
                    داخل الشركة
                  </option>

                  <option value="exited">
                    غادر
                  </option>

                  <option value="cancelled">
                    ملغاة
                  </option>
                </select>
              </div>

              <div className="table-wrapper">
                <table>
                  <thead>
                    <tr>
                      <th>
                        رقم الزيارة
                      </th>

                      <th>
                        التاريخ
                      </th>

                      <th>
                        الزائر
                      </th>

                      <th>
                        البريد الإلكتروني
                      </th>

                      <th>
                        الشركة
                      </th>

                      <th>
                        الحالة
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {filteredReportVisits.map(
                      (visit) => {
                        const status =
                          statusInfo(
                            visit.status,
                          );

                        return (
                          <tr
                            key={
                              visit.id
                            }
                          >
                            <td>
                              <strong>
                                {
                                  visit.visitNo
                                }
                              </strong>
                            </td>

                            <td>
                              {formatVisitDate(
                                visit.visitDate,
                              )}
                            </td>

                            <td>
                              {
                                visit.visitorName
                              }
                            </td>

                            <td>
                              {visit.visitorEmail || "—"}
                            </td>

                            <td>
                              {
                                visit.company
                              }
                            </td>

                            <td>
                              <span
                                className={
                                  "status " +
                                  status.className
                                }
                              >
                                <i>{status.icon}</i>
                                {status.label}
                              </span>
                            </td>
                          </tr>
                        );
                      },
                    )}

                    {filteredReportVisits.length ===
                      0 && (
                      <tr>
                        <td
                          colSpan={6}
                        >
                          لا توجد بيانات مطابقة
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
            )}

            {reportSection === "companies" && (
              <section className="visits-card">
                <div className="section-header"><div><h2>تقرير الشركات</h2><p>اسم الشركة والبريد الإلكتروني الذي أدخله الزائر</p></div></div>
                <div className="filters">
                  <div className="search-box"><span>⌕</span><input value={companyReportSearch} onChange={(event) => setCompanyReportSearch(event.target.value)} placeholder="ابحث باسم الشركة أو البريد الإلكتروني..." /></div>
                  <select value={companyReportEmailFilter} onChange={(event) => setCompanyReportEmailFilter(event.target.value as "all" | "withEmail" | "withoutEmail")} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}>
                    <option value="all">كل الشركات</option>
                    <option value="withEmail">شركات لديها إيميل</option>
                    <option value="withoutEmail">شركات بدون إيميل</option>
                  </select>
                </div>
                <div className="table-wrapper"><table><thead><tr><th>اسم الشركة</th><th>البريد الإلكتروني</th></tr></thead><tbody>
                  {filteredCompanyReport.map((company) => <tr key={company.id}><td><strong>{company.name}</strong></td><td>{company.email || "—"}</td></tr>)}
                  {filteredCompanyReport.length === 0 && <tr><td colSpan={2}>لا توجد شركات مطابقة</td></tr>}
                </tbody></table></div>
              </section>
            )}

            {reportSection === "mail" && (
              <section className="visits-card">
                <div className="section-header"><div><h2>تقرير البريد</h2><p>عرض وتصفية البريد الوارد والصادر</p></div></div>
                <div className="filters">
                  <div className="search-box"><span>⌕</span><input value={mailSearch} onChange={(event) => setMailSearch(event.target.value)} placeholder="ابحث برقم البريد أو الموضوع أو المسؤول..." /></div>
                  <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>من تاريخ <input type="date" value={reportMailFrom} onChange={(event) => setReportMailFrom(event.target.value)} aria-label="من تاريخ البريد" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                  <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>إلى تاريخ <input type="date" value={reportMailTo} onChange={(event) => setReportMailTo(event.target.value)} aria-label="إلى تاريخ البريد" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                  <select value={mailDirectionFilter} onChange={(event) => setMailDirectionFilter(event.target.value as "all" | MailDirection)} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}><option value="all">وارد وصادر</option><option value="incoming">وارد</option><option value="outgoing">صادر</option></select>
                  <select value={mailStatusFilter} onChange={(event) => setMailStatusFilter(event.target.value as "all" | MailStatusChoice)} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}><option value="all">كل الحالات</option><option value="followUp">قيد المتابعة</option><option value="waitingSignature">انتظار التوقيع</option><option value="processed">تم المعاملة</option><option value="preparing">قيد الإعداد</option><option value="other">أخرى</option></select>
                </div>
                <div className="table-wrapper"><table><thead><tr><th>رقم البريد</th><th>الموضوع</th><th>وارد / صادر</th><th>المسؤول</th><th>الإجراء</th><th>الحالة</th><th>التاريخ</th><th>الوقت</th></tr></thead><tbody>
                  {filteredMail.map((mail) => <tr key={mail.id}><td><strong>{mail.mailNo}</strong></td><td>{mail.subject}</td><td>{getMailDirectionLabel(mail.direction)}</td><td>{mail.responsible}</td><td>{mail.action}</td><td>{mail.status}</td><td>{formatVisitDate(mail.mailDate)}</td><td>{mail.mailTime}</td></tr>)}
                  {filteredMail.length === 0 && <tr><td colSpan={8}>لا توجد بيانات</td></tr>}
                </tbody></table></div>
              </section>
            )}

            {reportSection === "calls" && (
              <section className="visits-card">
                <div className="section-header"><div><h2>تقرير الاتصالات</h2><p>عرض وتصفية سجل الاتصالات</p></div></div>
                <div className="filters">
                  <div className="search-box"><span>⌕</span><input value={callSearch} onChange={(event) => setCallSearch(event.target.value)} placeholder="ابحث باسم المتصل أو الشركة أو الموضوع..." /></div>
                  <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>من تاريخ <input type="date" value={reportCallsFrom} onChange={(event) => setReportCallsFrom(event.target.value)} aria-label="من تاريخ الاتصالات" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                  <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>إلى تاريخ <input type="date" value={reportCallsTo} onChange={(event) => setReportCallsTo(event.target.value)} aria-label="إلى تاريخ الاتصالات" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                  <select value={callStatusFilter} onChange={(event) => setCallStatusFilter(event.target.value as "all" | CallStatus)} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}><option value="all">كل الحالات</option><option value="تحويل مكالمة">تحويل مكالمة</option><option value="تم">تم</option></select>
                </div>
                <div className="table-wrapper"><table><thead><tr><th>رقم الاتصال</th><th>اسم المتصل</th><th>الشركة</th><th>الرقم</th><th>الموضوع</th><th>الشخص المطلوب</th><th>تاريخ ووقت الاتصال</th><th>الحالة</th></tr></thead><tbody>
                  {filteredCalls.map((call) => <tr key={call.id}><td><strong>{call.callNo}</strong></td><td>{call.callerName}</td><td>{call.companyName}</td><td>{call.phone}</td><td>{call.subject}</td><td>{call.requestedEmployee}</td><td>{formatCallDateTime(call.createdAt)}</td><td>{(() => { const info = getCallStatusInfo(call.status); return <span className={"status " + info.className}><i>{info.icon}</i>{call.status}</span>; })()}</td></tr>)}
                  {filteredCalls.length === 0 && <tr><td colSpan={8}>لا توجد بيانات</td></tr>}
                </tbody></table></div>
              </section>
            )}

            {reportSection === "appointments" && (
              <section className="visits-card">
                <div className="section-header"><div><h2>تقرير المواعيد</h2><p>عرض وتصفية المواعيد</p></div></div>
                <div className="filters">
                  <div className="search-box"><span>⌕</span><input value={appointmentSearch} onChange={(event) => setAppointmentSearch(event.target.value)} placeholder="ابحث باسم الشخص أو المستضيف..." /></div>
                  <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>من تاريخ <input type="date" value={appointmentDateFrom} onChange={(event) => setAppointmentDateFrom(event.target.value)} aria-label="من تاريخ" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>إلى تاريخ <input type="date" value={appointmentDateTo} onChange={(event) => setAppointmentDateTo(event.target.value)} aria-label="إلى تاريخ" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} /></label>
                  <select value={appointmentStatusFilter} onChange={(event) => setAppointmentStatusFilter(event.target.value as "all" | AppointmentStatus)} style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px",background:"#fff"}}><option value="all">كل الحالات</option><option value="قيد الانتظار">قيد الانتظار</option><option value="تم الموعد">تم الموعد</option><option value="تأجيل الموعد">تأجيل الموعد</option><option value="إلغاء الموعد">إلغاء الموعد</option></select>
                </div>
                <div className="table-wrapper"><table><thead><tr><th>رقم الموعد</th><th>اسم الشخص</th><th>اليوم</th><th>التاريخ</th><th>وقت الموعد</th><th>اسم المستضيف</th><th>تاريخ الموعد</th><th>الحالة</th><th>التأجيل</th></tr></thead><tbody>
                  {filteredAppointments.map((appointment) => <tr key={appointment.id}><td><strong>{appointment.appointmentNo}</strong></td><td>{appointment.personName}</td><td>{appointment.dayName}</td><td>{formatVisitDate(appointment.entryDate)}</td><td>{appointment.appointmentTime}</td><td>{appointment.hostName}</td><td>{formatVisitDate(appointment.appointmentDate)}</td><td>{(() => { const info = getAppointmentStatusInfo(appointment.status); return <span className={"status " + info.className}><i>{info.icon}</i>{appointment.status}</span>; })()}</td><td>{appointment.status === "تأجيل الموعد" ? `${formatVisitDate(appointment.appointmentDate)} ${appointment.appointmentTime}` : "—"}</td></tr>)}
                  {filteredAppointments.length === 0 && <tr><td colSpan={9}>لا توجد بيانات</td></tr>}
                </tbody></table></div>
              </section>
            )}
          </>
        )}

        {/* =========================
            البريد
        ========================= */}

        {currentPage ===
          "mail" && (
          <>
            <header className="topbar">
              <div>
                <h1>
                  البريد
                </h1>

                <p>
                  إدارة البريد الوارد والصادر
                </p>
              </div>

              <div className="top-actions">
                <div className="current-date">
                  {mailItems.length} بريد
                </div>

                <button
                  type="button"
                  className="new-visit"
                  onClick={
                    openMailForm
                  }
                >
                  ＋ إضافة بريد
                </button>
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    قائمة البريد
                  </h2>

                  <p>
                    تسجيل ومتابعة البريد الوارد والصادر
                  </p>
                </div>
              </div>

              <div className="filters">
                <div className="search-box">
                  <span>
                    ⌕
                  </span>

                  <input
                    value={
                      mailSearch
                    }
                    onChange={(event) =>
                      setMailSearch(
                        event.target
                          .value,
                      )
                    }
                    placeholder="ابحث برقم البريد أو الموضوع أو المسؤول..."
                  />
                </div>

                <select
                  value={
                    mailDirectionFilter
                  }
                  onChange={(event) =>
                    setMailDirectionFilter(
                      event.target
                        .value as
                        | "all"
                        | MailDirection,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                    background:
                      "#fff",
                  }}
                >
                  <option value="all">
                    وارد وصادر
                  </option>

                  <option value="incoming">
                    وارد
                  </option>

                  <option value="outgoing">
                    صادر
                  </option>
                </select>

                <select
                  value={
                    mailStatusFilter
                  }
                  onChange={(event) =>
                    setMailStatusFilter(
                      event.target
                        .value as
                        | "all"
                        | MailStatusChoice,
                    )
                  }
                  style={{
                    height:
                      "42px",
                    border:
                      "1px solid #d1d5db",
                    borderRadius:
                      "9px",
                    padding:
                      "0 12px",
                    background:
                      "#fff",
                  }}
                >
                  <option value="all">
                    كل الحالات
                  </option>

                  <option value="followUp">
                    قيد المتابعة
                  </option>

                  <option value="waitingSignature">
                    انتظار التوقيع
                  </option>

                  <option value="processed">
                    تم المعاملة
                  </option>

                  <option value="preparing">
                    قيد الإعداد
                  </option>

                  <option value="other">
                    أخرى
                  </option>
                </select>

                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>
                  من تاريخ
                  <input type="date" value={mailDateFrom} onChange={(event) => setMailDateFrom(event.target.value)} aria-label="من تاريخ البريد" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} />
                </label>
                <label style={{display:"flex",alignItems:"center",gap:"7px",fontSize:"13px",fontWeight:700,color:"#374151"}}>
                  إلى تاريخ
                  <input type="date" value={mailDateTo} onChange={(event) => setMailDateTo(event.target.value)} aria-label="إلى تاريخ البريد" style={{height:"42px",border:"1px solid #d1d5db",borderRadius:"9px",padding:"0 12px"}} />
                </label>
              </div>

              {loadingMail && (
                <div className="empty-state">
                  جاري تحميل البريد...
                </div>
              )}

              {mailError && (
                <div className="empty-state">
                  {mailError}
                </div>
              )}

              {!loadingMail &&
                !mailError && (
                  <div className="table-wrapper">
                    <table>
                      <thead>
                        <tr>
                          <th>
                            رقم البريد
                          </th>

                          <th>
                            الموضوع
                          </th>

                          <th>
                            وارد / صادر
                          </th>

                          <th>
                            المسؤول
                          </th>

                          <th>
                            معلومات
                          </th>

                          <th>
                            الحالة
                          </th>

                          <th>
                            التاريخ
                          </th>

                          <th>
                            الوقت
                          </th>

                          <th>
                            المرفقات
                          </th>

                          <th>
                            الإجراء
                          </th>
                        </tr>
                      </thead>

                      <tbody>
                        {filteredMail.map(
                          (mail) => (
                            <tr
                              key={
                                mail.id
                              }
                            >
                              <td>
                                <strong>
                                  {
                                    mail.mailNo
                                  }
                                </strong>
                              </td>

                              <td>
                                <strong>
                                  {
                                    mail.subject
                                  }
                                </strong>
                              </td>

                              <td>
                                <span className={"status mail-direction " + getMailDirectionInfo(mail.direction).className}>
                                  <i>{getMailDirectionInfo(mail.direction).icon}</i>
                                  {getMailDirectionLabel(mail.direction)}
                                </span>
                              </td>

                              <td>
                                {
                                  mail.responsible
                                }
                              </td>

                              <td>
                                {mail.action ? (() => { const info = getMailActionInfo(mail.action); return <span className={"status action-badge " + info.className}><i>{info.icon}</i>{mail.action}</span>; })() : "—"}
                              </td>

                              <td>
                                <span className={"status " + getMailStatusInfo(mail.status).className}>
                                  <i>{getMailStatusInfo(mail.status).icon}</i>
                                  {mail.status || "—"}
                                </span>
                              </td>

                              <td>
                                {formatVisitDate(
                                  mail.mailDate,
                                )}
                              </td>

                              <td>
                                {
                                  mail.mailTime
                                }
                              </td>

                              <td>
                                {mail.attachments.length > 0 ? (
                                  <div
                                    style={{
                                      display: "flex",
                                      flexDirection: "column",
                                      gap: "6px",
                                    }}
                                  >
                                    <span
                                      style={{
                                        fontSize: "12px",
                                        fontWeight: 700,
                                        color: "#475569",
                                      }}
                                    >
                                      {mail.attachments.length} ملف
                                    </span>
                                    {mail.attachments.slice(0, 2).map((attachment) => (
                                      <button
                                        key={attachment.id}
                                        type="button"
                                        className="action-button details"
                                        onClick={() => void openMailAttachment(attachment)}
                                        style={{ whiteSpace: "nowrap" }}
                                      >
                                        فتح المرفق
                                      </button>
                                    ))}
                                  </div>
                                ) : (
                                  "—"
                                )}
                              </td>

                              <td>
                                <div
                                  style={{
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    gap: "7px",
                                    flexWrap: "wrap",
                                  }}
                                >
                                  <button
                                    type="button"
                                    className="action-button details"
                                    onClick={() =>
                                      openEditMail(
                                        mail,
                                      )
                                    }
                                  >
                                    تفاصيل
                                  </button>

                                  <button
                                    type="button"
                                    className="action-button"
                                    onClick={() => void deleteMail(mail)}
                                    style={{
                                      background: "#dc2626",
                                      color: "#fff",
                                    }}
                                  >
                                    حذف
                                  </button>
                                </div>
                              </td>
                            </tr>
                          ),
                        )}

                        {filteredMail.length ===
                          0 && (
                          <tr>
                            <td
                              colSpan={10}
                            >
                              لا توجد رسائل مطابقة
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
            </section>
          </>
        )}

        {/* =========================
            الإعدادات
        ========================= */}

        {currentPage ===
          "settings" && (
          <>
            <header className="topbar">
              <div>
                <h1>
                  الإعدادات
                </h1>

                <p>
                  تخصيص مظهر نظام إدارة الزيارات
                </p>
              </div>

              <div className="top-actions">
                <div className="current-date">
                  اللون الحالي
                </div>

                <div
                  className="color-preview-small"
                  style={{
                    background:
                      accentColor,
                  }}
                />
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    لون النظام
                  </h2>

                  <p>
                    اختر اللون الرئيسي للنظام
                  </p>
                </div>
              </div>

              <div className="settings-content">
                <div className="color-grid">
                  {[
                    {
                      name: "أزرق",
                      color: "#2563eb",
                    },
                    {
                      name: "نيلي",
                      color: "#4f46e5",
                    },
                    {
                      name: "أخضر",
                      color: "#059669",
                    },
                    {
                      name: "تركوازي",
                      color: "#0891b2",
                    },
                    {
                      name: "بنفسجي",
                      color: "#7c3aed",
                    },
                    {
                      name: "برتقالي",
                      color: "#ea580c",
                    },
                    {
                      name: "أحمر",
                      color: "#dc2626",
                    },
                  ].map((item) => (
                    <button
                      key={
                        item.color
                      }
                      type="button"
                      className={
                        "color-option " +
                        (accentColor ===
                        item.color
                          ? "selected"
                          : "")
                      }
                      onClick={() =>
                        setAccentColor(
                          item.color,
                        )
                      }
                    >
                      <span
                        style={{
                          background:
                            item.color,
                        }}
                      />

                      <strong>
                        {
                          item.name
                        }
                      </strong>
                    </button>
                  ))}
                </div>

                <div className="custom-color">
                  <label>
                    لون مخصص
                  </label>

                  <div className="custom-color-row">
                    <input
                      type="color"
                      value={
                        accentColor
                      }
                      onChange={(
                        event,
                      ) =>
                        setAccentColor(
                          event.target
                            .value,
                        )
                      }
                    />

                    <div>
                      <strong>
                        {accentColor.toUpperCase()}
                      </strong>

                      <span>
                        اختر أي لون تريده
                      </span>
                    </div>
                  </div>
                </div>

                <div className="settings-preview">
                  <div className="settings-preview-title">
                    معاينة مباشرة
                  </div>

                  <div className="settings-preview-box">
                    <button
                      type="button"
                      className="new-visit"
                    >
                      زر تجريبي
                    </button>

                    <button
                      type="button"
                      className="action-button details"
                    >
                      إجراء
                    </button>

                    <span
                      className="preview-primary-text"
                      style={{
                        color:
                          accentColor,
                      }}
                    >
                      النص الرئيسي
                    </span>
                  </div>
                </div>

                <button
                  type="button"
                  className="view-all settings-reset"
                  onClick={() =>
                    setAccentColor(
                      "#2563eb",
                    )
                  }
                >
                  إعادة اللون الافتراضي
                </button>
              </div>
            </section>
          </>
        )}

        {/* =========================
            تسجيل زيارة جديدة
        ========================= */}

        {currentPage ===
          "newVisit" && (
          <>
            <header className="topbar">
              <div>
                <h1>
                  تسجيل زيارة جديدة
                </h1>

                <p>
                  إضافة زيارة جديدة للنظام
                </p>
              </div>

              <div className="top-actions">
                <button
                  type="button"
                  className="view-all"
                  onClick={() =>
                    setCurrentPage(
                      "home",
                    )
                  }
                >
                  ← العودة للرئيسية
                </button>
              </div>
            </header>

            <section className="visits-card">
              <div className="section-header">
                <div>
                  <h2>
                    بيانات الزيارة
                  </h2>

                  <p>
                    ابدأ باختيار الشركة المرتبطة بالزائر
                  </p>
                </div>

                <div className="current-date">
                  الخطوة 1 من 4
                </div>
              </div>

              {!selectedCompany && (
                <>
                  <div className="filters">
                    <div className="search-box">
                      <span>
                        ⌕
                      </span>

                      <input
                        value={
                          companySearch
                        }
                        onChange={(event) =>
                          setCompanySearch(
                            event.target
                              .value,
                          )
                        }
                        placeholder="ابحث عن الشركة..."
                      />
                    </div>

                    <button
                      type="button"
                      className="new-visit"
                      onClick={() =>
                        openAddCompany(
                          true,
                        )
                      }
                    >
                      ＋ إضافة شركة جديدة
                    </button>

                    <button
                      type="button"
                      className="filter-button"
                      onClick={
                        selectPersonal
                      }
                    >
                      شخصي
                    </button>
                  </div>

                  {loadingCompanies && (
                    <div className="empty-state">
                      جاري تحميل الشركات...
                    </div>
                  )}

                  {companyError && (
                    <div className="empty-state">
                      {companyError}
                    </div>
                  )}

                  {!loadingCompanies &&
                    !companyError && (
                      <div className="table-wrapper">
                        <table>
                          <thead>
                            <tr>
                              <th>
                                الشركة
                              </th>

                              <th>
                                الحالة
                              </th>

                              <th>
                                الإجراء
                              </th>
                            </tr>
                          </thead>

                          <tbody>
                            {filteredCompanies.map(
                              (
                                company,
                              ) => (
                                <tr
                                  key={
                                    company.id
                                  }
                                >
                                  <td>
                                    <button
                                      type="button"
                                      className="company-name-button"
                                      onClick={() =>
                                        selectCompany(
                                          company,
                                        )
                                      }
                                    >
                                      {
                                        company.name
                                      }
                                    </button>
                                  </td>

                                  <td>
                                    <span className="status inside">
                                      <i />
                                      شركة مسجلة
                                    </span>
                                  </td>

                                  <td>
                                    <button
                                      type="button"
                                      className="action-button check-in"
                                      onClick={() =>
                                        selectCompany(
                                          company,
                                        )
                                      }
                                    >
                                      اختيار الشركة
                                    </button>
                                  </td>
                                </tr>
                              ),
                            )}

                            {filteredCompanies.length ===
                              0 && (
                              <tr>
                                <td
                                  colSpan={
                                    3
                                  }
                                >
                                  لا توجد شركة مطابقة للبحث
                                  <div
                                    style={{
                                      marginTop:
                                        "10px",
                                    }}
                                  >
                                    <button
                                      type="button"
                                      className="new-visit"
                                      onClick={() =>
                                        openAddCompany(
                                          true,
                                        )
                                      }
                                    >
                                      ＋ إضافة شركة جديدة
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    )}
                </>
              )}

              {selectedCompany && (
                <>
                  <div className="section-header selected-company">
                    <div>
                      <h2>
                        الشركة المختارة
                      </h2>

                      <p>
                        {
                          selectedCompany
                        }
                      </p>
                    </div>

                    <button
                      type="button"
                      className="view-all"
                      onClick={
                        changeCompany
                      }
                    >
                      تغيير الشركة
                    </button>
                  </div>

                  {selectedCompanyId !==
                    null && (
                    <>
                      <div className="section-header">
                        <div>
                          <h2>
                            الزوار السابقون
                          </h2>

                          <p>
                            اختر زائرًا سبق له زيارة هذه الشركة
                          </p>
                        </div>

                        <button
                          type="button"
                          className="new-visit"
                          onClick={
                            startNewVisitor
                          }
                        >
                          ＋ زائر جديد
                        </button>
                      </div>

                      {loadingVisitors && (
                        <div className="empty-state">
                          جاري تحميل الزوار السابقين...
                        </div>
                      )}

                      {visitorError && (
                        <div className="empty-state">
                          {visitorError}
                        </div>
                      )}

                      {!loadingVisitors &&
                        !visitorError &&
                        !showNewVisitor &&
                        !selectedVisitor && (
                          <>
                            {previousVisitors.length >
                            0 ? (
                              <div className="table-wrapper">
                                <table>
                                  <thead>
                                    <tr>
                                      <th>
                                        اسم الزائر
                                      </th>

                                      <th>
                                        الهاتف
                                      </th>

                                      <th>
                                        الإيميل
                                      </th>

                                      <th>
                                        الإجراء
                                      </th>
                                    </tr>
                                  </thead>

                                  <tbody>
                                    {previousVisitors.map(
                                      (
                                        visitor,
                                      ) => (
                                        <tr
                                          key={
                                            visitor.id
                                          }
                                        >
                                          <td>
                                            <strong>
                                              {
                                                visitor.visitorName
                                              }
                                            </strong>
                                          </td>

                                          <td>
                                            {
                                              visitor.visitorPhone ||
                                              "—"
                                            }
                                          </td>

                                          <td>
                                            {
                                              visitor.visitorEmail ||
                                              "—"
                                            }
                                          </td>

                                          <td>
                                            <button
                                              type="button"
                                              className="action-button check-in"
                                              onClick={() =>
                                                selectPreviousVisitor(
                                                  visitor,
                                                )
                                              }
                                            >
                                              اختيار الزائر
                                            </button>
                                          </td>
                                        </tr>
                                      ),
                                    )}
                                  </tbody>
                                </table>
                              </div>
                            ) : (
                              <div className="empty-state">
                                لا يوجد زوار سابقون لهذه الشركة
                              </div>
                            )}
                          </>
                        )}

                      {showNewVisitor && (
                        <VisitorForm
                          visitorName={
                            visitorName
                          }
                          setVisitorName={
                            setVisitorName
                          }
                          visitorPhone={
                            visitorPhone
                          }
                          setVisitorPhone={
                            setVisitorPhone
                          }
                          visitorEmail={
                            visitorEmail
                          }
                          setVisitorEmail={
                            setVisitorEmail
                          }
                          requestedEmployee={
                            requestedEmployee
                          }
                          setRequestedEmployee={
                            setRequestedEmployee
                          }
                          visitReason={
                            visitReason
                          }
                          setVisitReason={
                            setVisitReason
                          }
                          formError={
                            formError
                          }
                          savingVisit={
                            savingVisit
                          }
                          formInputStyle={
                            formInputStyle
                          }
                          onCancel={() =>
                            setShowNewVisitor(
                              false,
                            )
                          }
                          onSave={() =>
                            void saveNewVisit()
                          }
                        />
                      )}

                      {selectedVisitor && (
                        <div className="section-header selected-company">
                          <div>
                            <h2>
                              الزائر المختار
                            </h2>

                            <p>
                              {
                                selectedVisitor.visitorName
                              }

                              {" • "}

                              {selectedVisitor.visitorPhone ||
                                "بدون هاتف"}
                            </p>
                          </div>

                          <button
                            type="button"
                            className="new-visit"
                            onClick={() => {
                              setShowNewVisitor(
                                true,
                              );

                              setFormError(
                                "",
                              );
                            }}
                          >
                            متابعة ←
                          </button>
                        </div>
                      )}
                    </>
                  )}

                  {selectedCompanyId ===
                    null &&
                    showNewVisitor && (
                    <VisitorForm
                      visitorName={
                        visitorName
                      }
                      setVisitorName={
                        setVisitorName
                      }
                      visitorPhone={
                        visitorPhone
                      }
                      setVisitorPhone={
                        setVisitorPhone
                      }
                      visitorEmail={
                        visitorEmail
                      }
                      setVisitorEmail={
                        setVisitorEmail
                      }
                      requestedEmployee={
                        requestedEmployee
                      }
                      setRequestedEmployee={
                        setRequestedEmployee
                      }
                      visitReason={
                        visitReason
                      }
                      setVisitReason={
                        setVisitReason
                      }
                      formError={
                        formError
                      }
                      savingVisit={
                        savingVisit
                      }
                      formInputStyle={
                        formInputStyle
                      }
                      onCancel={() =>
                        setShowNewVisitor(
                          false,
                        )
                      }
                      onSave={() =>
                        void saveNewVisit()
                      }
                    />
                  )}
                </>
              )}
            </section>
          </>
        )}
      </main>

      {/* =========================
          نموذج الاتصال
      ========================= */}

      {showCallForm && (
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.45)",display:"flex",alignItems:"center",justifyContent:"center",padding:"20px",zIndex:1000}} onMouseDown={(event) => { if (event.target === event.currentTarget) closeCallForm(); }}>
          <div style={{width:"100%",maxWidth:"720px",maxHeight:"90vh",overflowY:"auto",background:"#fff",borderRadius:"16px",boxShadow:"0 20px 60px rgba(0,0,0,0.18)",padding:"26px"}} dir="rtl">
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:"16px",marginBottom:"24px"}}><div><h2 style={{margin:0,fontSize:"20px",color:"#111827"}}>{callEditingId !== null ? "تعديل الاتصال" : "إضافة اتصال جديد"}</h2><p style={{margin:"7px 0 0",color:"#6b7280",fontSize:"13px"}}>إدخال بيانات الاتصال بشكل يدوي</p></div><button type="button" className="view-all" onClick={closeCallForm} disabled={savingCall}>× إغلاق</button></div>
            {callEditingId !== null && <div style={{marginBottom:"18px",padding:"11px 14px",borderRadius:"9px",background:"#eff6ff",border:"1px solid #bfdbfe",color:"#1d4ed8",fontSize:"13px",fontWeight:700}}>رقم الاتصال ثابت: {callItems.find((item) => item.id === callEditingId)?.callNo}</div>}
            <div className="form-grid">
              <div className="form-field"><label>اسم متصل</label><input style={formInputStyle} value={callerName} onChange={(event) => setCallerName(event.target.value)} /></div>
              <div className="form-field"><label>الشركة</label><input style={formInputStyle} value={callCompanyName} onChange={(event) => setCallCompanyName(event.target.value)} /></div>
              <div className="form-field"><label>الرقم</label><input style={formInputStyle} value={callPhone} onChange={(event) => setCallPhone(event.target.value)} /></div>
              <div className="form-field"><label>الموضوع</label><input style={formInputStyle} value={callSubject} onChange={(event) => setCallSubject(event.target.value)} /></div>
              <div className="form-field"><label>الشخص المطلوب</label><input style={formInputStyle} value={callRequestedEmployee} onChange={(event) => setCallRequestedEmployee(event.target.value)} /></div>
              <div className="form-field"><label>الحالة</label><select style={formInputStyle} value={callStatus} onChange={(event) => setCallStatus(event.target.value as CallStatus)}><option value="تحويل مكالمة">تحويل مكالمة</option><option value="تم">تم</option></select></div>
            </div>
            {callFormError && <div style={{marginTop:"16px",padding:"10px 12px",borderRadius:"9px",background:"#fef2f2",color:"#b91c1c"}}>{callFormError}</div>}
            <div className="form-actions"><button type="button" className="new-visit" onClick={() => void saveCall()} disabled={savingCall}>{savingCall ? "جاري الحفظ..." : "حفظ"}</button><button type="button" className="view-all" onClick={closeCallForm} disabled={savingCall}>إلغاء</button></div>
          </div>
        </div>
      )}

      {/* =========================
          نموذج الموعد
      ========================= */}

      {showAppointmentForm && (
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.45)",display:"flex",alignItems:"center",justifyContent:"center",padding:"20px",zIndex:1000}} onMouseDown={(event) => { if (event.target === event.currentTarget) closeAppointmentForm(); }}>
          <div style={{width:"100%",maxWidth:"780px",maxHeight:"90vh",overflowY:"auto",background:"#fff",borderRadius:"16px",boxShadow:"0 20px 60px rgba(0,0,0,0.18)",padding:"26px"}} dir="rtl">
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:"16px",marginBottom:"24px"}}><div><h2 style={{margin:0,fontSize:"20px",color:"#111827"}}>{appointmentEditingId !== null ? "تعديل الموعد" : "إضافة موعد جديد"}</h2><p style={{margin:"7px 0 0",color:"#6b7280",fontSize:"13px"}}>إدخال بيانات الموعد بشكل يدوي</p></div><button type="button" className="view-all" onClick={closeAppointmentForm} disabled={savingAppointment}>× إغلاق</button></div>
            {appointmentEditingId !== null && <div style={{marginBottom:"18px",padding:"11px 14px",borderRadius:"9px",background:"#eff6ff",border:"1px solid #bfdbfe",color:"#1d4ed8",fontSize:"13px",fontWeight:700}}>رقم الموعد ثابت: {appointmentItems.find((item) => item.id === appointmentEditingId)?.appointmentNo}</div>}
            <div className="form-grid">
              <div className="form-field"><label>اسم الشخص</label><input style={formInputStyle} value={appointmentPersonName} onChange={(event) => setAppointmentPersonName(event.target.value)} /></div>
              <div className="form-field"><label>اليوم</label><input style={formInputStyle} value={appointmentDayName} onChange={(event) => setAppointmentDayName(event.target.value)} placeholder="مثال: الأحد" /></div>
              <div className="form-field"><label>التاريخ</label><input type="date" style={formInputStyle} value={appointmentEntryDate} onChange={(event) => setAppointmentEntryDate(event.target.value)} /></div>
              <div className="form-field"><label>وقت الموعد (توقيت الأردن)</label><input type="time" style={formInputStyle} value={appointmentTime} onChange={(event) => setAppointmentTime(event.target.value)} /></div>
              <div className="form-field"><label>اسم المستضيف</label><input style={formInputStyle} value={appointmentHostName} onChange={(event) => setAppointmentHostName(event.target.value)} /></div>
              <div className="form-field"><label>تاريخ الموعد</label><input type="date" style={formInputStyle} value={appointmentDate} onChange={(event) => setAppointmentDate(event.target.value)} /></div>
              <div className="form-field"><label>الحالة</label><select style={formInputStyle} value={appointmentStatus} onChange={(event) => { const value = event.target.value as AppointmentStatus; setAppointmentStatus(value); if (value === "تأجيل الموعد") { setAppointmentNewDate(appointmentDate); setAppointmentNewTime(appointmentTime); } }}><option value="قيد الانتظار">قيد الانتظار</option><option value="تم الموعد">تم الموعد</option><option value="تأجيل الموعد">تأجيل الموعد</option><option value="إلغاء الموعد">إلغاء الموعد</option></select></div>
              {appointmentStatus === "تأجيل الموعد" && <div className="form-field"><label>التاريخ الجديد</label><input type="date" style={formInputStyle} value={appointmentNewDate} onChange={(event) => setAppointmentNewDate(event.target.value)} /></div>}
              {appointmentStatus === "تأجيل الموعد" && <div className="form-field"><label>الوقت الجديد</label><input type="time" style={formInputStyle} value={appointmentNewTime} onChange={(event) => setAppointmentNewTime(event.target.value)} /></div>}
            </div>
            {appointmentFormError && <div style={{marginTop:"16px",padding:"10px 12px",borderRadius:"9px",background:"#fef2f2",color:"#b91c1c"}}>{appointmentFormError}</div>}
            <div className="form-actions"><button type="button" className="new-visit" onClick={() => void saveAppointment()} disabled={savingAppointment}>{savingAppointment ? "جاري الحفظ..." : "حفظ"}</button><button type="button" className="view-all" onClick={closeAppointmentForm} disabled={savingAppointment}>إلغاء</button></div>
          </div>
        </div>
      )}

      {/* =========================
          نموذج البريد
      ========================= */}

      {showMailForm && (
        <div
          style={{
            position:
              "fixed",
            inset: 0,
            background:
              "rgba(15, 23, 42, 0.45)",
            display:
              "flex",
            alignItems:
              "center",
            justifyContent:
              "center",
            padding:
              "20px",
            zIndex:
              1000,
          }}
          onMouseDown={(
            event,
          ) => {
            if (
              event.target ===
              event.currentTarget
            ) {
              closeMailForm();
            }
          }}
        >
          <div
            style={{
              width:
                "100%",
              maxWidth:
                "720px",
              maxHeight:
                "90vh",
              overflowY:
                "auto",
              background:
                "#fff",
              borderRadius:
                "16px",
              boxShadow:
                "0 20px 60px rgba(0,0,0,0.18)",
              padding:
                "26px",
            }}
            dir="rtl"
          >
            <div
              style={{
                display:
                  "flex",
                alignItems:
                  "center",
                justifyContent:
                  "space-between",
                gap:
                  "16px",
                marginBottom:
                  "24px",
              }}
            >
              <div>
                <h2
                  style={{
                    margin:
                      0,
                    fontSize:
                      "20px",
                    color:
                      "#111827",
                  }}
                >
                  {mailEditingId !==
                  null
                    ? "تعديل البريد"
                    : "إضافة بريد جديد"}
                </h2>

                <p
                  style={{
                    margin:
                      "7px 0 0",
                    color:
                      "#6b7280",
                    fontSize:
                      "13px",
                  }}
                >
                  {mailEditingId !==
                  null
                    ? "تعديل بيانات البريد الحالي"
                    : "إدخال بيانات البريد بشكل يدوي"}
                </p>
              </div>

              <button
                type="button"
                className="view-all"
                onClick={
                  closeMailForm
                }
                disabled={
                  savingMail
                }
              >
                × إغلاق
              </button>
            </div>

            {mailEditingId !==
              null && (
              <div
                style={{
                  marginBottom:
                    "18px",
                  padding:
                    "11px 14px",
                  borderRadius:
                    "9px",
                  background:
                    "#eff6ff",
                  border:
                    "1px solid #bfdbfe",
                  color:
                    "#1d4ed8",
                  fontSize:
                    "13px",
                  fontWeight:
                    700,
                }}
              >
                رقم البريد ثابت:
                {" "}
                {
                  mailItems.find(
                    (item) =>
                      item.id ===
                      mailEditingId,
                  )?.mailNo
                }
              </div>
            )}

            <div
              style={{
                display:
                  "grid",
                gridTemplateColumns:
                  "repeat(2, minmax(0, 1fr))",
                gap:
                  "18px",
              }}
            >
              <div
                style={{
                  gridColumn:
                    "1 / -1",
                }}
              >
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  الموضوع *
                </label>

                <input
                  type="text"
                  value={
                    mailSubject
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailSubject(
                      event.target
                        .value,
                    )
                  }
                  placeholder="أدخل موضوع البريد"
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                />
              </div>

              <div>
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  وارد / صادر *
                </label>

                <select
                  value={
                    mailDirection
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailDirection(
                      event.target
                        .value as MailDirection,
                    )
                  }
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                >
                  <option value="incoming">
                    وارد
                  </option>

                  <option value="outgoing">
                    صادر
                  </option>
                </select>
              </div>

              <div>
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  المسؤول *
                </label>

                <input
                  type="text"
                  value={
                    mailResponsible
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailResponsible(
                      event.target
                        .value,
                    )
                  }
                  placeholder="اسم المسؤول"
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                />
              </div>

              <div>
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  الإجراء *
                </label>

                <select
                  value={
                    mailActionChoice
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailActionChoice(
                      event.target
                        .value as MailActionChoice,
                    )
                  }
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                >
                  <option value="signed">
                    تم توقيع
                  </option>

                  <option value="dataEntry">
                    قيد الإدخال
                  </option>

                  <option value="other">
                    أخرى
                  </option>
                </select>
              </div>

              {mailActionChoice ===
                "other" && (
                <div>
                  <label
                    style={{
                      display:
                        "block",
                      marginBottom:
                        "8px",
                      fontSize:
                        "13px",
                      fontWeight:
                        700,
                      color:
                        "#374151",
                    }}
                  >
                    اكتب الإجراء *
                  </label>

                  <input
                    type="text"
                    value={
                      mailActionOther
                    }
                    onChange={(
                      event,
                    ) =>
                      setMailActionOther(
                        event.target
                          .value,
                      )
                    }
                    placeholder="اكتب الإجراء"
                    style={
                      formInputStyle
                    }
                    disabled={
                      savingMail
                    }
                  />
                </div>
              )}

              <div>
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  الحالة *
                </label>

                <select
                  value={
                    mailStatusChoice
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailStatusChoice(
                      event.target
                        .value as MailStatusChoice,
                    )
                  }
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                >
                  <option value="followUp">
                    قيد المتابعة
                  </option>

                  <option value="waitingSignature">
                    انتظار التوقيع
                  </option>

                  <option value="processed">
                    تم المعاملة
                  </option>

                  <option value="preparing">
                    قيد الإعداد
                  </option>

                  <option value="other">
                    أخرى
                  </option>
                </select>
              </div>

              {mailStatusChoice ===
                "other" && (
                <div>
                  <label
                    style={{
                      display:
                        "block",
                      marginBottom:
                        "8px",
                      fontSize:
                        "13px",
                      fontWeight:
                        700,
                      color:
                        "#374151",
                    }}
                  >
                    اكتب الحالة *
                  </label>

                  <input
                    type="text"
                    value={
                      mailStatusOther
                    }
                    onChange={(
                      event,
                    ) =>
                      setMailStatusOther(
                        event.target
                          .value,
                      )
                    }
                    placeholder="اكتب الحالة"
                    style={
                      formInputStyle
                    }
                    disabled={
                      savingMail
                    }
                  />
                </div>
              )}

              <div>
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  التاريخ *
                </label>

                <input
                  type="date"
                  value={
                    mailDate
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailDate(
                      event.target
                        .value,
                    )
                  }
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                />
              </div>

              <div>
                <label
                  style={{
                    display:
                      "block",
                    marginBottom:
                      "8px",
                    fontSize:
                      "13px",
                    fontWeight:
                      700,
                    color:
                      "#374151",
                  }}
                >
                  الوقت *
                </label>

                <input
                  type="time"
                  value={
                    mailTime
                  }
                  onChange={(
                    event,
                  ) =>
                    setMailTime(
                      event.target
                        .value,
                    )
                  }
                  style={
                    formInputStyle
                  }
                  disabled={
                    savingMail
                  }
                />
              </div>
            </div>

            <div
              style={{
                marginTop: "20px",
                padding: "16px",
                border: "1px dashed #cbd5e1",
                borderRadius: "12px",
                background: "#f8fafc",
              }}
            >
              <label
                style={{
                  display: "block",
                  marginBottom: "8px",
                  fontSize: "13px",
                  fontWeight: 800,
                  color: "#374151",
                }}
              >
                المرفقات
              </label>

              <input
                type="file"
                multiple
                onChange={(event) =>
                  setMailSelectedFiles(
                    Array.from(event.target.files ?? []),
                  )
                }
                style={{
                  width: "100%",
                  padding: "10px",
                  border: "1px solid #d1d5db",
                  borderRadius: "9px",
                  background: "#fff",
                }}
                disabled={savingMail}
              />

              <div
                style={{
                  marginTop: "8px",
                  fontSize: "12px",
                  color: "#64748b",
                }}
              >
                يمكنك إرفاق أكثر من ملف مع البريد.
              </div>

              {mailSelectedFiles.length > 0 && (
                <div
                  style={{
                    marginTop: "12px",
                    display: "flex",
                    flexDirection: "column",
                    gap: "6px",
                  }}
                >
                  {mailSelectedFiles.map((file, index) => (
                    <div
                      key={`${file.name}-${file.size}-${index}`}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        gap: "12px",
                        padding: "9px 11px",
                        background: "#fff",
                        borderRadius: "8px",
                        border: "1px solid #e2e8f0",
                        fontSize: "12px",
                      }}
                    >
                      <span
                        style={{
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                        }}
                      >
                        {file.name}
                      </span>
                      <span
                        style={{
                          color: "#64748b",
                          flexShrink: 0,
                        }}
                      >
                        {(file.size / 1024).toFixed(0)} KB
                      </span>
                    </div>
                  ))}
                </div>
              )}

              {mailEditingId !== null && (
                <div
                  style={{
                    marginTop: "14px",
                    paddingTop: "12px",
                    borderTop: "1px solid #e2e8f0",
                  }}
                >
                  <div
                    style={{
                      fontSize: "12px",
                      fontWeight: 800,
                      color: "#475569",
                      marginBottom: "8px",
                    }}
                  >
                    المرفقات المحفوظة
                  </div>

                  {(mailItems.find((item) => item.id === mailEditingId)?.attachments ?? []).length > 0 ? (
                    <div
                      style={{
                        display: "flex",
                        flexDirection: "column",
                        gap: "7px",
                      }}
                    >
                      {(mailItems.find((item) => item.id === mailEditingId)?.attachments ?? []).map((attachment) => (
                        <div
                          key={attachment.id}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "space-between",
                            gap: "12px",
                            padding: "9px 11px",
                            background: "#fff",
                            borderRadius: "8px",
                            border: "1px solid #e2e8f0",
                          }}
                        >
                          <span
                            style={{
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                              fontSize: "12px",
                            }}
                          >
                            {attachment.fileName}
                          </span>

                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: "7px",
                              flexShrink: 0,
                            }}
                          >
                            <button
                              type="button"
                              className="action-button details"
                              onClick={() => void openMailAttachment(attachment)}
                            >
                              فتح
                            </button>

                            <button
                              type="button"
                              className="action-button"
                              onClick={() => void deleteMailAttachment(attachment)}
                              disabled={savingMail}
                              style={{
                                background: "#dc2626",
                                color: "#fff",
                              }}
                            >
                              حذف
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div
                      style={{
                        fontSize: "12px",
                        color: "#94a3b8",
                      }}
                    >
                      لا توجد مرفقات محفوظة.
                    </div>
                  )}
                </div>
              )}
            </div>

            {mailFormError && (
              <div
                style={{
                  marginTop:
                    "16px",
                  padding:
                    "11px 14px",
                  borderRadius:
                    "9px",
                  background:
                    "#fef2f2",
                  border:
                    "1px solid #fecaca",
                  color:
                    "#b91c1c",
                  fontSize:
                    "13px",
                  fontWeight:
                    600,
                }}
              >
                {
                  mailFormError
                }
              </div>
            )}

            <div
              style={{
                display:
                  "flex",
                gap:
                  "10px",
                marginTop:
                  "24px",
              }}
            >
              <button
                type="button"
                className="view-all"
                disabled={
                  savingMail
                }
                onClick={
                  closeMailForm
                }
              >
                إلغاء
              </button>

              <button
                type="button"
                className="new-visit"
                disabled={
                  savingMail
                }
                onClick={() =>
                  void saveMail()
                }
              >
                {savingMail
                  ? "جاري الحفظ..."
                  : mailEditingId !==
                      null
                    ? "حفظ التعديل"
                    : "حفظ البريد"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* =========================
          نموذج الشركة
      ========================= */}

      {showCompanyForm && (
        <div
          style={{
            position:
              "fixed",
            inset: 0,
            background:
              "rgba(15, 23, 42, 0.45)",
            display:
              "flex",
            alignItems:
              "center",
            justifyContent:
              "center",
            padding:
              "20px",
            zIndex:
              1000,
          }}
          onMouseDown={(
            event,
          ) => {
            if (
              event.target ===
              event.currentTarget
            ) {
              closeCompanyForm();
            }
          }}
        >
          <div
            style={{
              width:
                "100%",
              maxWidth:
                "520px",
              background:
                "#fff",
              borderRadius:
                "16px",
              boxShadow:
                "0 20px 60px rgba(0,0,0,0.18)",
              padding:
                "26px",
            }}
            dir="rtl"
          >
            <div
              style={{
                display:
                  "flex",
                alignItems:
                  "center",
                justifyContent:
                  "space-between",
                gap:
                  "16px",
                marginBottom:
                  "24px",
              }}
            >
              <div>
                <h2
                  style={{
                    margin:
                      0,
                    fontSize:
                      "20px",
                    color:
                      "#111827",
                  }}
                >
                  {companyFormMode ===
                  "add"
                    ? "إضافة شركة جديدة"
                    : "تعديل الشركة"}
                </h2>

                <p
                  style={{
                    margin:
                      "7px 0 0",
                    color:
                      "#6b7280",
                    fontSize:
                      "13px",
                  }}
                >
                  {companyFormMode ===
                  "add"
                    ? companyFormFromVisit
                      ? "إضافة الشركة ومتابعة تسجيل الزيارة"
                      : "أدخل بيانات الشركة الجديدة"
                    : "تعديل بيانات الشركة"}
                </p>
              </div>

              <button
                type="button"
                className="view-all"
                onClick={
                  closeCompanyForm
                }
                disabled={
                  savingCompany
                }
              >
                × إغلاق
              </button>
            </div>

            <div>
              <label
                style={{
                  display:
                    "block",
                  marginBottom:
                    "8px",
                  fontSize:
                    "13px",
                  fontWeight:
                    700,
                  color:
                    "#374151",
                }}
              >
                اسم الشركة *
              </label>

              <input
                type="text"
                value={
                  companyFormName
                }
                onChange={(
                  event,
                ) =>
                  setCompanyFormName(
                    event.target
                      .value,
                  )
                }
                placeholder="أدخل اسم الشركة"
                autoFocus
                style={
                  formInputStyle
                }
                disabled={
                  savingCompany
                }
              />
            </div>

            <label
              style={{
                display:
                  "flex",
                alignItems:
                  "center",
                gap:
                  "10px",
                marginTop:
                  "18px",
                fontSize:
                  "13px",
                fontWeight:
                  700,
                color:
                  "#374151",
                cursor:
                  "pointer",
              }}
            >
              <input
                type="checkbox"
                checked={
                  companyFormActive
                }
                onChange={(
                  event,
                ) =>
                  setCompanyFormActive(
                    event
                      .target
                      .checked,
                  )
                }
                disabled={
                  savingCompany
                }
              />

              الشركة فعالة
            </label>

            {companyFormError && (
              <div
                style={{
                  marginTop:
                    "16px",
                  padding:
                    "11px 14px",
                  borderRadius:
                    "9px",
                  background:
                    "#fef2f2",
                  border:
                    "1px solid #fecaca",
                  color:
                    "#b91c1c",
                  fontSize:
                    "13px",
                  fontWeight:
                    600,
                }}
              >
                {
                  companyFormError
                }
              </div>
            )}

            <div
              style={{
                display:
                  "flex",
                gap:
                  "10px",
                marginTop:
                  "24px",
              }}
            >
              <button
                type="button"
                className="view-all"
                disabled={
                  savingCompany
                }
                onClick={
                  closeCompanyForm
                }
              >
                إلغاء
              </button>

              <button
                type="button"
                className="new-visit"
                disabled={
                  savingCompany
                }
                onClick={() =>
                  void saveCompany()
                }
              >
                {savingCompany
                  ? "جاري الحفظ..."
                  : companyFormMode ===
                      "add"
                    ? "حفظ الشركة"
                    : "حفظ التعديل"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* =========================
   نموذج بيانات الزائر
========================= */

type VisitorFormProps = {
  visitorName: string;
  setVisitorName: (
    value: string,
  ) => void;

  visitorPhone: string;
  setVisitorPhone: (
    value: string,
  ) => void;

  visitorEmail: string;
  setVisitorEmail: (
    value: string,
  ) => void;

  requestedEmployee: string;
  setRequestedEmployee: (
    value: string,
  ) => void;

  visitReason: string;
  setVisitReason: (
    value: string,
  ) => void;

  formError: string;
  savingVisit: boolean;

  formInputStyle: CSSProperties;

  onCancel: () => void;
  onSave: () => void;
};

function VisitorForm({
  visitorName,
  setVisitorName,
  visitorPhone,
  setVisitorPhone,
  visitorEmail,
  setVisitorEmail,
  requestedEmployee,
  setRequestedEmployee,
  visitReason,
  setVisitReason,
  formError,
  savingVisit,
  formInputStyle,
  onCancel,
  onSave,
}: VisitorFormProps) {
  return (
    <div
      style={{
        marginTop:
          "20px",
        padding:
          "24px",
        border:
          "1px solid #e5e7eb",
        borderRadius:
          "14px",
        background:
          "#fff",
      }}
    >
      <div
        style={{
          display:
            "grid",
          gridTemplateColumns:
            "repeat(2, minmax(0, 1fr))",
          gap:
            "18px",
        }}
      >
        <div>
          <label
            style={{
              display:
                "block",
              marginBottom:
                "8px",
              fontSize:
                "13px",
              fontWeight:
                700,
              color:
                "#374151",
            }}
          >
            اسم الزائر *
          </label>

          <input
            type="text"
            value={
              visitorName
            }
            onChange={(
              event,
            ) =>
              setVisitorName(
                event.target
                  .value,
              )
            }
            placeholder="أدخل اسم الزائر"
            style={
              formInputStyle
            }
          />
        </div>

        <div>
          <label
            style={{
              display:
                "block",
              marginBottom:
                "8px",
              fontSize:
                "13px",
              fontWeight:
                700,
              color:
                "#374151",
            }}
          >
            رقم الهاتف
          </label>

          <input
            type="text"
            value={
              visitorPhone
            }
            onChange={(
              event,
            ) =>
              setVisitorPhone(
                event.target
                  .value,
              )
            }
            placeholder="أدخل رقم الهاتف"
            style={
              formInputStyle
            }
          />
        </div>

        <div>
          <label
            style={{
              display:
                "block",
              marginBottom:
                "8px",
              fontSize:
                "13px",
              fontWeight:
                700,
              color:
                "#374151",
            }}
          >
            البريد الإلكتروني
          </label>

          <input
            type="email"
            value={
              visitorEmail
            }
            onChange={(
              event,
            ) =>
              setVisitorEmail(
                event.target
                  .value,
              )
            }
            placeholder="أدخل البريد الإلكتروني"
            style={
              formInputStyle
            }
          />
        </div>

        <div>
          <label
            style={{
              display:
                "block",
              marginBottom:
                "8px",
              fontSize:
                "13px",
              fontWeight:
                700,
              color:
                "#374151",
            }}
          >
            الشخص المطلوب *
          </label>

          <input
            type="text"
            value={
              requestedEmployee
            }
            onChange={(
              event,
            ) =>
              setRequestedEmployee(
                event.target
                  .value,
              )
            }
            placeholder="اسم الموظف المطلوب"
            style={
              formInputStyle
            }
          />
        </div>

        <div
          style={{
            gridColumn:
              "1 / -1",
          }}
        >
          <label
            style={{
              display:
                "block",
              marginBottom:
                "8px",
              fontSize:
                "13px",
              fontWeight:
                700,
              color:
                "#374151",
            }}
          >
            الغرض من الزيارة *
          </label>

          <input
            type="text"
            value={
              visitReason
            }
            onChange={(
              event,
            ) =>
              setVisitReason(
                event.target
                  .value,
              )
            }
            placeholder="أدخل الغرض من الزيارة"
            style={
              formInputStyle
            }
          />
        </div>
      </div>

      {formError && (
        <div
          style={{
            marginTop:
              "16px",
            padding:
              "11px 14px",
            borderRadius:
              "9px",
            background:
              "#fef2f2",
            border:
              "1px solid #fecaca",
            color:
              "#b91c1c",
            fontSize:
              "13px",
            fontWeight:
              600,
          }}
        >
          {
            formError
          }
        </div>
      )}

      <div
        style={{
          display:
            "flex",
          justifyContent:
            "flex-start",
          gap:
            "10px",
          marginTop:
            "24px",
        }}
      >
        <button
          type="button"
          className="view-all"
          disabled={
            savingVisit
          }
          onClick={
            onCancel
          }
        >
          إلغاء
        </button>

        <button
          type="button"
          className="new-visit"
          disabled={
            savingVisit
          }
          onClick={
            onSave
          }
        >
          {savingVisit
            ? "جاري الحفظ..."
            : "حفظ ومتابعة ←"}
        </button>
      </div>
    </div>
  );
}

export default App;

