import React, { Fragment, useCallback, useEffect, useMemo, useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import QR_BASE64 from './qrData.js'

const DEFAULT_MEMBERS = [
  'Dr. B. Anusha Balaram', 'Chandu Krishna', 'K. V. T. Ramesh', 'Rama Devi',
  'K. V. T. Chandra Mouli', 'Tarak', 'Naveen', 'Nageswara Rao', 'Rajesh Kumar V',
  'Manikanta Teja K', 'Ajay Kumar P', 'Rajesh V', 'Naveen Kumar K',
  'Dhilleswara Rao D', 'Santosh Kumar G', 'Venkateswarlu T', 'Ravikumar G',
  'Kalyani B', 'Viswa Sai Dhanush K', 'Balaram', 'V. Chandra Krishna',
  'Sateesh A', 'Kalyan Bogi', 'Nagraju G', 'Gopal', 'Dr. John', 'Ramesh T',
  'S. Udaya Bhaskar Rao', 'Nageswararao K', 'Ajay Kumar', 'Kalkinadh Gudla',
  'Santosh Kumar P', 'Rajasekhar V', 'Santosh Reddi', 'Uday K',
  'Tharak Pydisetty', 'Bhavani', 'Manikanta', 'V. Santosh Kumar', 'Anusha',
]

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
]

const STATUS_CONFIG = {
  present: { label: 'Present', color: '#166534', bg: '#dcfce7' },
  absent: { label: 'Absent', color: '#991b1b', bg: '#fee2e2' },
  substitute: { label: 'Substitute', color: '#92400e', bg: '#fef3c7' },
}

const STORAGE_KEYS = {
  members: 'bni-dashboard-members-v2',
  data: 'bni-dashboard-month-data-v2',
  visitors: 'bni-dashboard-visitors-v2',
  fridayMode: 'bni-dashboard-friday-mode-session',
  paymentFilter: 'bni-dashboard-payment-filter-session',
  paymentStatus: 'bni-dashboard-payment-status-session',
  expenses: 'bni-dashboard-expenses-v1',
}

function safeJsonRead(key, fallback, storage = window.localStorage) {
  try {
    const value = storage.getItem(key)
    return value ? JSON.parse(value) : fallback
  } catch {
    return fallback
  }
}

function getFridays(year, month) {
  const fridays = []
  const d = new Date(year, month, 1)
  while (d.getDay() !== 5) d.setDate(d.getDate() + 1)
  while (d.getMonth() === month) {
    fridays.push(new Date(d))
    d.setDate(d.getDate() + 7)
  }
  return fridays
}

function startOfDay(date) {
  const next = new Date(date)
  next.setHours(0, 0, 0, 0)
  return next
}

function sameDate(a, b) {
  return a.getFullYear() === b.getFullYear()
    && a.getMonth() === b.getMonth()
    && a.getDate() === b.getDate()
}

function fridayMeta(date) {
  const year = date.getFullYear()
  const month = date.getMonth()
  const fridays = getFridays(year, month)
  return {
    date: new Date(date),
    year,
    month,
    index: Math.max(0, fridays.findIndex((friday) => sameDate(friday, date))),
  }
}

function getVisibleFridays(year, month, mode) {
  const monthFridays = getFridays(year, month)
  if (mode === 'all') return monthFridays.map(fridayMeta)

  const today = startOfDay(new Date())
  const upcoming = monthFridays.find((date) => startOfDay(date) >= today)

  if (upcoming) return [fridayMeta(upcoming)]

  const nextMonth = month === 11 ? 0 : month + 1
  const nextYear = month === 11 ? year + 1 : year
  return [fridayMeta(getFridays(nextYear, nextMonth)[0])]
}

function fmtDate(date) {
  return `${date.getDate()} ${MONTHS[date.getMonth()].slice(0, 3)} ${date.getFullYear()}`
}

function calcAmount(status, payMode, totalFridays , fridayIndex) {
  // Weekly members unchanged
  if (payMode === 'weekly') {
    if (status === 'absent') return 1200
    if (status === 'substitute') return 1000
    return 800
  }

  // Monthly members
  const baseFee = totalFridays >= 5 ? 4300 : 3000

  // First Friday
  if (fridayIndex === 0) {
    if (status === 'absent') return baseFee + 400
    if (status === 'substitute') return baseFee + 200
    return baseFee
  }

  // Remaining FridaysS
  if (status === 'absent') return 400
  if (status === 'substitute') return 200
  return 0
}

function initFriday(status = 'present', payMode = 'weekly', totalFridays = 4, fridayIndex = 0) {
  return {
    status,
    amount: calcAmount(status, payMode, totalFridays, fridayIndex),
    payMethod: 'cash',
    paid: false,
  }
}

function initMemberRow(fridayCount) {
  return {
    payMode: 'weekly',
    fridays: Array.from({ length: fridayCount }, (_, index) => initFriday('present', 'weekly', fridayCount, index)),
  }
}

function normalizeMemberRow(row, fridayCount) {
  const payMode = row?.payMode === 'monthly' ? 'monthly' : 'weekly'
  const fridays = Array.from({ length: fridayCount }, (_, index) => {
    const existing = row?.fridays?.[index]
    if (!existing) return initFriday('present', payMode, fridayCount, index)
    return {
      status: existing.status && STATUS_CONFIG[existing.status] ? existing.status : 'present',
      amount: Number(existing.amount) || calcAmount(existing.status, payMode, fridayCount,index),
      payMethod: existing.payMethod === 'upi' ? 'upi' : 'cash',
      paid: Boolean(existing.paid),
    }
  })
  return { payMode, fridays }
}

function normalizeMonthData(raw, fridayCount, memberCount) {
  return {
    members: Array.from({ length: memberCount }, (_, index) => {
      return normalizeMemberRow(raw?.members?.[index], fridayCount)
    }),
    rent: {
      amount: Number(raw?.rent?.amount) || 13000,
      status: raw?.rent?.status === 'paid' ? 'paid' : 'not_paid',
    },
  }
}

function initMonthData(year, month, memberCount) {
  return normalizeMonthData(null, getFridays(year, month).length, memberCount)
}

function initials(name) {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase())
    .join('')
}

function makeId() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`
}

const css = `
  *, *::before, *::after { box-sizing: border-box; }
  body { margin: 0; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f5f7fb; color: #172033; }
  button, input, select { font: inherit; }
  button { cursor: pointer; }

  .login-wrap { min-height: 100vh; display: grid; place-items: center; padding: 24px; background: #eef2f7; }
  .login-card { width: min(420px, 100%); background: #ffffff; border: 1px solid #d9e0ea; border-radius: 12px; padding: 36px; box-shadow: 0 24px 70px rgba(23, 32, 51, 0.14); }
  .login-logo { font-size: 2.4rem; font-weight: 850; letter-spacing: 0.08em; color: #b91c1c; text-align: center; }
  .login-sub { text-align: center; margin: 8px 0 28px; color: #516070; font-weight: 700; font-size: 0.84rem; text-transform: uppercase; letter-spacing: 0.08em; }
  .field { display: grid; gap: 7px; margin-bottom: 14px; }
  .field label, .form-label { color: #516070; font-size: 0.78rem; font-weight: 800; letter-spacing: 0.04em; text-transform: uppercase; }
  .input, .select { width: 100%; border: 1px solid #cfd8e3; border-radius: 8px; background: #ffffff; color: #172033; padding: 10px 12px; outline: none; min-height: 40px; }
  .input:focus, .select:focus { border-color: #315a94; box-shadow: 0 0 0 3px rgba(49, 90, 148, 0.14); }
  .login-error, .form-error { color: #b91c1c; font-size: 0.86rem; font-weight: 700; margin-top: 10px; }

  .app-shell { min-height: 100vh; }
  .topbar { position: sticky; top: 0; z-index: 40; display: flex; justify-content: space-between; align-items: center; gap: 18px; min-height: 64px; padding: 0 24px; background: #172033; color: #ffffff; box-shadow: 0 1px 12px rgba(23, 32, 51, 0.18); }
  .brand { display: flex; align-items: center; gap: 14px; min-width: 0; }
  .brand-mark { color: #ffffff; font-weight: 900; letter-spacing: 0.12em; font-size: 1.35rem; }
  .brand-sub { color: #aab6c8; font-size: 0.84rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .nav { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
  .nav-btn, .ghost-btn, .primary-btn, .danger-btn, .muted-btn { border: 1px solid transparent; border-radius: 8px; min-height: 38px; padding: 8px 13px; font-size: 0.86rem; font-weight: 800; transition: background 0.15s, border-color 0.15s, color 0.15s; }
  .nav-btn { background: transparent; color: #dce5f2; border-color: transparent; }
  .nav-btn:hover, .nav-btn.active { background: #26344d; color: #ffffff; }
  .ghost-btn { background: transparent; color: #ffffff; border-color: #52627a; }
  .ghost-btn:hover { background: #26344d; }
  .primary-btn { background: #1f5f9f; color: #ffffff; }
  .primary-btn:hover { background: #174b80; }
  .danger-btn { background: #b91c1c; color: #ffffff; }
  .danger-btn:hover { background: #991b1b; }
  .muted-btn { background: #eef2f7; color: #172033; border-color: #d7e0ea; }
  .muted-btn:hover { background: #e4ebf4; }

  .main { width: min(1760px, 100%); margin: 0 auto; padding: 22px 18px 72px; }
  .page-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; margin-bottom: 18px; }
  .page-title { margin: 0; font-size: clamp(1.45rem, 2vw, 2rem); letter-spacing: 0; }
  .page-copy { margin: 5px 0 0; color: #627083; font-size: 0.92rem; }

  .toolbar { display: flex; align-items: end; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
  .control { display: grid; gap: 6px; min-width: 130px; }
  .control.compact { min-width: 104px; }
  .segmented { display: inline-flex; align-items: center; border: 1px solid #cfd8e3; border-radius: 9px; background: #ffffff; overflow: hidden; min-height: 40px; }
  .segment { border: 0; background: transparent; color: #516070; min-height: 38px; padding: 8px 14px; font-size: 0.84rem; font-weight: 800; }
  .segment.active { background: #1f5f9f; color: #ffffff; }
  .status-note { align-self: center; color: #627083; font-size: 0.88rem; font-weight: 700; padding: 9px 0; }

  .quick-actions { display: flex; align-items: end; gap: 8px; flex-wrap: wrap; }
  .summary-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin-bottom: 16px; }
  .summary-card { background: #ffffff; border: 1px solid #d9e0ea; border-radius: 10px; padding: 16px; box-shadow: 0 8px 28px rgba(23, 32, 51, 0.05); }
  .summary-label { color: #627083; font-size: 0.78rem; font-weight: 850; letter-spacing: 0.04em; text-transform: uppercase; }
  .summary-value { margin-top: 8px; color: #172033; font-size: clamp(1.35rem, 2vw, 1.85rem); font-weight: 900; line-height: 1.1; }
  .summary-value.success { color: #166534; }
  .summary-value.warning { color: #991b1b; }

  .panel { background: #ffffff; border: 1px solid #d9e0ea; border-radius: 10px; padding: 16px; margin-bottom: 16px; box-shadow: 0 8px 28px rgba(23, 32, 51, 0.05); }
  .rent-panel { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
  .panel-title { margin: 0; font-size: 1rem; font-weight: 850; color: #172033; }
  .panel-copy { margin: 3px 0 0; color: #627083; font-size: 0.86rem; }
  .rent-fields { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; margin-left: auto; }
  .amount-input { width: 118px; text-align: right; font-weight: 800; }
  .badge { display: inline-flex; align-items: center; border-radius: 999px; padding: 6px 11px; font-size: 0.8rem; font-weight: 850; }
  .badge.ok { background: #dcfce7; color: #166534; }
  .badge.warn { background: #fee2e2; color: #991b1b; }

  .legend { display: flex; gap: 12px; flex-wrap: wrap; color: #516070; font-size: 0.82rem; }
  .legend strong { color: #172033; }
  .legend-item { display: inline-flex; gap: 6px; align-items: center; }
  .legend-dot { width: 10px; height: 10px; border-radius: 99px; display: inline-block; }

  .table-wrap { overflow-x: auto; background: #ffffff; border: 1px solid #d9e0ea; border-radius: 10px; box-shadow: 0 14px 38px rgba(23, 32, 51, 0.08); }
  table { width: 100%; min-width: max-content; border-collapse: separate; border-spacing: 0; font-size: 0.84rem; }
  th, td { border-right: 1px solid #e4eaf2; border-bottom: 1px solid #e4eaf2; padding: 9px 8px; text-align: center; vertical-align: middle; }
  thead th { position: sticky; z-index: 10; top: 0; background: #172033; color: #ffffff; font-weight: 850; white-space: nowrap; }
  thead tr:nth-child(2) th { top: 39px; background: #26344d; }
  th:first-child, td:first-child { border-left: 0; }
  tbody tr:nth-child(even) td { background: #f9fbfd; }
  tbody tr:hover td { background: #eef4fb; }
  .name-head { background: #172033; }
  .name-cell { background: inherit; text-align: left; min-width: 210px; font-weight: 800; color: #172033; }
  .sno-cell { color: #627083; min-width: 46px; }
  .summary-row td { background: #172033 !important; color: #ffffff; font-weight: 900; }
  .total-cell { min-width: 112px; font-weight: 900; color: #166534; background: #f0fdf4 !important; }
  .paid-total-cell { min-width: 100px; font-weight: 900; color: #1f5f9f; background: #eff6ff !important; }
  .mini-select { min-width: 112px; border: 1px solid #cfd8e3; border-radius: 7px; background: #ffffff; color: #172033; padding: 6px 8px; outline: none; }
  .mini-input { width: 88px; border: 1px solid #cfd8e3; border-radius: 7px; background: #ffffff; color: #172033; padding: 6px 8px; text-align: right; outline: none; font-weight: 800; }
  .payment-cell { display: grid; gap: 8px; justify-items: start; min-width: 126px; }
  .paid-check { display: inline-flex; align-items: center; gap: 7px; color: #344154; font-size: 0.8rem; font-weight: 800; }
  .paid-check input { width: 15px; height: 15px; }
  .paymode-toggle { display: inline-flex; border: 1px solid #cfd8e3; border-radius: 8px; overflow: hidden; background: #ffffff; }
  .paymode-btn { border: 0; background: #ffffff; color: #516070; padding: 6px 10px; font-size: 0.78rem; font-weight: 900; min-width: 40px; }
  .paymode-btn.active { background: #1f5f9f; color: #ffffff; }
  .empty-state { padding: 32px; color: #627083; text-align: center; }

  .form-grid { display: grid; grid-template-columns: minmax(220px, 1fr) auto; gap: 12px; align-items: end; }
  .visitor-form-grid { grid-template-columns: minmax(220px, 1fr) 120px 130px 90px auto; }
  .management-list { display: grid; gap: 10px; }
  .member-row, .visitor-row { display: grid; grid-template-columns: 48px minmax(180px, 1fr) auto; gap: 12px; align-items: center; padding: 12px; border: 1px solid #d9e0ea; border-radius: 9px; background: #ffffff; }
  .avatar { width: 36px; height: 36px; display: grid; place-items: center; border-radius: 999px; background: #e8eef7; color: #315a94; font-weight: 900; font-size: 0.8rem; }
  .row-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
  .visitor-row { grid-template-columns: minmax(180px, 1fr) 120px 130px 90px auto; }
  .qr-section { margin-top: 22px; text-align: center; background: #172033; border-radius: 12px; padding: 28px 20px; color: #ffffff; }
  .qr-title { margin: 0; font-size: 1rem; font-weight: 900; letter-spacing: 0.06em; }
  .qr-sub { margin: 7px 0 18px; color: #aab6c8; font-size: 0.88rem; }
  .qr-img-wrap { display: inline-block; padding: 8px; border-radius: 12px; background: #ffffff; }
  .qr-img { display: block; max-width: min(320px, 82vw); height: auto; border-radius: 8px; }
  .qr-name { margin-top: 16px; font-weight: 850; }

  @media (max-width: 820px) {
    .topbar { align-items: flex-start; flex-direction: column; padding: 14px; }
    .nav { justify-content: flex-start; }
    .main { padding: 16px 10px 60px; }
    .page-head, .rent-panel { align-items: stretch; flex-direction: column; }
    .rent-fields { margin-left: 0; }
    .form-grid, .visitor-form-grid, .member-row, .visitor-row { grid-template-columns: 1fr; }
    .avatar { display: none; }
    .row-actions { justify-content: flex-start; }
    .segmented { width: 100%; }
    .segment { flex: 1; }
    .toolbar .control { min-width: calc(50% - 8px); }
    .summary-grid { grid-template-columns: 1fr; }
    .quick-actions { width: 100%; }
    .quick-actions .muted-btn { flex: 1; }
  }
`

export default function App() {
  const [loggedIn, setLoggedIn] = useState(false)
  const [user, setUser] = useState('')
  const [pass, setPass] = useState('')
  const [loginErr, setLoginErr] = useState('')
  const [page, setPage] = useState('dashboard')

  const now = new Date()
  const [year, setYear] = useState(now.getFullYear())
  const [month, setMonth] = useState(now.getMonth())

  const [members, setMembers] = useState(() => safeJsonRead(STORAGE_KEYS.members, DEFAULT_MEMBERS))
  const [allData, setAllData] = useState(() => safeJsonRead(STORAGE_KEYS.data, {}))
  const [visitors, setVisitors] = useState(() => safeJsonRead(STORAGE_KEYS.visitors, []))
  const [fridayMode, setFridayMode] = useState(() => {
    return safeJsonRead(STORAGE_KEYS.fridayMode, 'current', window.sessionStorage)
  })
  const [paymentFilter, setPaymentFilter] = useState(() => {
    return safeJsonRead(STORAGE_KEYS.paymentFilter, 'all', window.sessionStorage)
  })
  const [paymentStatus, setPaymentStatus] = useState(() => {
    return safeJsonRead(STORAGE_KEYS.paymentStatus, 'all', window.sessionStorage)
  })

  const [newMemberName, setNewMemberName] = useState('')
  const [memberError, setMemberError] = useState('')
  const [editingIndex, setEditingIndex] = useState(null)
  const [editingName, setEditingName] = useState('')

  const [visitorForm, setVisitorForm] = useState({
    name: '',
    friday: '',
    referredBy: '',
    amount: 300,
    paymentMethod: 'cash',
    paid: false,
  })
  const [visitorError, setVisitorError] = useState('')

  const [expenses, setExpenses] = useState(() => safeJsonRead(STORAGE_KEYS.expenses, {}))

  const getExpensesFor = (targetYear, targetMonth) => {
    return {
      weeklyRent: expenses?.[targetYear]?.[targetMonth]?.weeklyRent ?? 0,
      notablePrizes: expenses?.[targetYear]?.[targetMonth]?.notablePrizes ?? 0,
      rosterSheet: expenses?.[targetYear]?.[targetMonth]?.rosterSheet ?? 0,
    }
  }

  const setExpensesFor = (targetYear, targetMonth, newExpenses) => {
    setExpenses((prev) => ({
      ...prev,
      [targetYear]: {
        ...(prev?.[targetYear] ?? {}),
        [targetMonth]: {
          ...(prev?.[targetYear]?.[targetMonth] ?? {}),
          ...newExpenses,
        },
      },
    }))
  }

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.expenses, JSON.stringify(expenses))
  }, [expenses])

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.members, JSON.stringify(members))
  }, [members])

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.data, JSON.stringify(allData))
  }, [allData])

  useEffect(() => {
    localStorage.setItem(STORAGE_KEYS.visitors, JSON.stringify(visitors))
  }, [visitors])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.fridayMode, JSON.stringify(fridayMode))
  }, [fridayMode])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.paymentFilter, JSON.stringify(paymentFilter))
  }, [paymentFilter])

  useEffect(() => {
    sessionStorage.setItem(STORAGE_KEYS.paymentStatus, JSON.stringify(paymentStatus))
  }, [paymentStatus])

  // Auto-select Friday for visitor form
  useEffect(() => {
    const today = startOfDay(new Date())
    let selectedFriday = null

    // Check current month first
    let currentFridays = getFridays(year, month)
    for (const friday of currentFridays) {
      const fri = startOfDay(friday)
      if (fri >= today) {
        selectedFriday = fri
        break
      }
    }

    // If no Fridays left in current month, check next month
    if (!selectedFriday) {
      let nextMonth = month === 11 ? 0 : month + 1
      let nextYear = month === 11 ? year + 1 : year
      const nextMonthFridays = getFridays(nextYear, nextYear)
      if (nextMonthFridays.length > 0) {
        selectedFriday = startOfDay(nextMonthFridays[0])
      }
    }

    if (selectedFriday) {
      setVisitorForm((prev) => ({ ...prev, friday: selectedFriday.toISOString() }))
    }
  }, [year, month])

  const getMonthDataFor = useCallback((targetYear, targetMonth) => {
    const fridayCount = getFridays(targetYear, targetMonth).length
    return normalizeMonthData(allData?.[targetYear]?.[targetMonth], fridayCount, members.length)
  }, [allData, members.length])

  const setMonthDataFor = useCallback((targetYear, targetMonth, updater) => {
    setAllData((prev) => {
      const fridayCount = getFridays(targetYear, targetMonth).length
      const current = normalizeMonthData(prev?.[targetYear]?.[targetMonth], fridayCount, members.length)
      const next = typeof updater === 'function' ? updater(current) : updater
      return {
        ...prev,
        [targetYear]: {
          ...(prev[targetYear] ?? {}),
          [targetMonth]: next,
        },
      }
    })
  }, [members.length])

  const monthData = useMemo(() => getMonthDataFor(year, month), [getMonthDataFor, year, month])
  const visibleFridays = useMemo(() => getVisibleFridays(year, month, fridayMode), [year, month, fridayMode])
  const contextMeta = fridayMode === 'all'
    ? { year, month }
    : { year: visibleFridays[0]?.year ?? year, month: visibleFridays[0]?.month ?? month }
  const contextMonthData = getMonthDataFor(contextMeta.year, contextMeta.month)

  const years = []
  for (let y = now.getFullYear() - 1; y <= now.getFullYear() + 2; y++) years.push(y)

  const handleLogin = () => {
    if (user.trim() === 'Chandu' && pass === 'admin@1234') {
      setLoggedIn(true)
      setLoginErr('')
    } else {
      setLoginErr('Invalid username or password.')
    }
  }

  const updatePayMode = (memberIndex, mode) => {
    setMonthDataFor(contextMeta.year, contextMeta.month, (prev) => {
      const membersRows = prev.members.map((row, index) => {
        if (index !== memberIndex) return row
        const updatedFridays = row.fridays.map((friday, fridayIndex) => ({
          ...friday,
          amount: calcAmount(friday.status, mode, row.fridays.length, fridayIndex),
        }))
        return { ...row, payMode: mode, fridays: updatedFridays }
      })
      return { ...prev, members: membersRows }
    })
  }

  const updateFridayField = (memberIndex, friday, field, value) => {
    setMonthDataFor(friday.year, friday.month, (prev) => {
      const membersRows = prev.members.map((row, rowIndex) => {
        if (rowIndex !== memberIndex) return row
        const fridays = row.fridays.map((entry, fridayIndex) => {
          if (fridayIndex !== friday.index) return entry
          const updated = { ...entry, [field]: value }
          if (field === 'status') {
            updated.amount = calcAmount(value, row.payMode, row.fridays.length, fridayIndex)
          }
          return updated
        })
        return { ...row, fridays }
      })
      return { ...prev, members: membersRows }
    })
  }

  const updateRent = (field, value) => {
    setMonthDataFor(year, month, (prev) => ({
      ...prev,
      rent: { ...prev.rent, [field]: value },
    }))
  }

  const getFridayEntry = (memberIndex, friday) => {
    return getMonthDataFor(friday.year, friday.month).members[memberIndex]?.fridays[friday.index]
      ?? initFriday('present', 'weekly', getFridays(friday.year, friday.month).length, friday.index)
  }

  const rowTotals = useMemo(() => {
    return members.map((_, memberIndex) => {
      return visibleFridays.reduce((total, friday) => {
        const entry = getFridayEntry(memberIndex, friday)
        return total + (Number(entry.amount) || 0)
      }, 0)
    })
  }, [members, visibleFridays, allData])

  const rowPaidTotals = useMemo(() => {
    return members.map((_, memberIndex) => {
      return visibleFridays.reduce((total, friday) => {
        const entry = getFridayEntry(memberIndex, friday)
        return total + (entry.paid ? Number(entry.amount) || 0 : 0)
      }, 0)
    })
  }, [members, visibleFridays, allData])

  const memberPaymentSummaries = useMemo(() => {
    return members.map((_, memberIndex) => {
      const entries = visibleFridays.map((friday) => getFridayEntry(memberIndex, friday))
      const totalAmount = entries.reduce((sum, entry) => sum + (Number(entry.amount) || 0), 0)
      const collectedAmount = entries.reduce((sum, entry) => {
        return sum + (entry.paid ? Number(entry.amount) || 0 : 0)
      }, 0)
      const outstandingAmount = Math.max(0, totalAmount - collectedAmount)
      const hasEntries = entries.length > 0
      const isPaid = hasEntries && entries.every((entry) => entry.paid)
      const isUnpaid = !isPaid

      return {
        totalAmount,
        collectedAmount,
        outstandingAmount,
        isPaid,
        isUnpaid,
      }
    })
  }, [members, visibleFridays, allData])

  const displayedMemberIndexes = useMemo(() => {
    return members
      .map((_, index) => index)
      .filter((index) => {
        const typeMatches = paymentFilter === 'all' || contextMonthData.members[index]?.payMode === paymentFilter
        const statusMatches = paymentStatus === 'all'
          || (paymentStatus === 'paid' && memberPaymentSummaries[index]?.isPaid)
          || (paymentStatus === 'unpaid' && memberPaymentSummaries[index]?.isUnpaid)
        return typeMatches && statusMatches
      })
  }, [members, paymentFilter, paymentStatus, contextMonthData, memberPaymentSummaries])

  const visitorCollectedTotal = visitors.reduce((sum, visitor) => {
    return sum + (visitor.paid ? Number(visitor.amount) || 0 : 0)
  }, 0)
  const visitorOutstandingTotal = visitors.reduce((sum, visitor) => {
    return sum + (!visitor.paid ? Number(visitor.amount) || 0 : 0)
  }, 0)

  const summaryStats = useMemo(() => {
    const totalMembersPaid = memberPaymentSummaries.filter((summary) => summary.isPaid).length
    const totalMembersUnpaid = memberPaymentSummaries.filter((summary) => summary.isUnpaid).length
    const memberCollection = memberPaymentSummaries.reduce((sum, summary) => {
      return sum + summary.collectedAmount
    }, 0)
    const visitorCollection = visitorCollectedTotal
    const totalAmountCollected = memberCollection + visitorCollection
    const outstandingAmount = memberPaymentSummaries.reduce((sum, summary) => {
      return sum + summary.outstandingAmount
    }, 0) + visitorOutstandingTotal

    // Grand Total Payable: sum of all members' totalAmount
    const grandTotalPayable = memberPaymentSummaries.reduce((sum, summary) => {
      return sum + summary.totalAmount
    }, 0)

    const currentExpenses = getExpensesFor(year, month)
    const totalExpenses = currentExpenses.weeklyRent + currentExpenses.notablePrizes + currentExpenses.rosterSheet
    const netCollection = Math.max(0, grandTotalPayable - totalExpenses)

    return {
      totalMembersPaid,
      totalMembersUnpaid,
      totalAmountCollected,
      outstandingAmount,
      memberCollection,
      visitorCollection,
      grandTotalPayable,
      totalExpenses,
      netCollection,
    }
  }, [memberPaymentSummaries, visitorCollectedTotal, visitorOutstandingTotal, year, month, expenses])

  const visibleGrandTotal = displayedMemberIndexes.reduce((sum, index) => sum + rowTotals[index], 0)
  const visiblePaidTotal = displayedMemberIndexes.reduce((sum, index) => sum + rowPaidTotals[index], 0)

  const exportExcel = () => {
    const wb = XLSX.utils.book_new()
    const row1 = ['#', 'Member Name', 'Pay Mode']
    const row2 = ['', '', '']

    visibleFridays.forEach((friday, index) => {
      row1.push(`Friday ${index + 1} (${fmtDate(friday.date)})`, '', '', '')
      row2.push('Attendance', 'Amount (Rs.)', 'Payment Method', 'Paid')
    })
    row1.push('Total Payable (Rs.)', 'Due (Rs.)', 'Total Paid (Rs.)')
    row2.push('', '', '')

    const rows = [row1, row2]
    let grandPayable = 0
    let grandPaid = 0

    displayedMemberIndexes.forEach((memberIndex, rowIndex) => {
      const payMode = contextMonthData.members[memberIndex]?.payMode ?? 'weekly'
      const row = [rowIndex + 1, members[memberIndex], payMode === 'weekly' ? 'Weekly' : 'Monthly']
      let totalPayable = 0
      let totalPaid = 0

      visibleFridays.forEach((friday) => {
        const entry = getFridayEntry(memberIndex, friday)
        const amount = Number(entry.amount) || 0
        totalPayable += amount
        if (entry.paid) totalPaid += amount
        row.push(
          STATUS_CONFIG[entry.status]?.label ?? 'Present',
          amount,
          entry.payMethod.toUpperCase(),
          entry.paid ? 'Paid' : 'Unpaid',
        )
      })

      grandPayable += totalPayable
      grandPaid += totalPaid
      const due = Math.max(0, totalPayable - totalPaid)
      row.push(totalPayable, due, totalPaid)
      rows.push(row)
    })

    rows.push(['', 'GRAND TOTAL', '', ...visibleFridays.flatMap(() => ['', '', '', '']), grandPayable, Math.max(0, grandPayable - grandPaid), grandPaid])
    rows.push([])
    rows.push([
      'VENUE RENT',
      '',
      'Amount (Rs.)',
      monthData.rent.amount,
      'Status',
      monthData.rent.status === 'paid' ? 'Paid' : 'Not Paid',
    ])

    // Sheet 1: Dashboard
    const dashboardWs = XLSX.utils.aoa_to_sheet(rows)
    dashboardWs['!cols'] = [
      { wch: 4 },
      { wch: 28 },
      { wch: 12 },
      ...visibleFridays.flatMap(() => [{ wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 10 }]),
      { wch: 18 },
      { wch: 14 },
      { wch: 16 },
    ]
    const merges = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
      { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
    ]
    visibleFridays.forEach((_, index) => {
      const c = 3 + index * 4
      merges.push({ s: { r: 0, c }, e: { r: 0, c: c + 3 } })
    })
    dashboardWs['!merges'] = merges
    XLSX.utils.book_append_sheet(wb, dashboardWs, 'Dashboard')

    // Sheet 2: Visitors (filtered for current month/year)
    const visitorRows = [
      ['Visitor Name', 'Friday', 'Referred By', 'Amount', 'Payment Method', 'Paid Status'],
    ]
    const filteredVisitors = visitors.filter((visitor) => {
      if (!visitor.friday) return false
      const d = new Date(visitor.friday)
      return d.getFullYear() === year && d.getMonth() === month
    })
    filteredVisitors.forEach((visitor) => {
      const visitorFriday = visitor.friday ? fmtDate(new Date(visitor.friday)) : ''
      visitorRows.push([
        visitor.name,
        visitorFriday,
        visitor.referredBy || '',
        Number(visitor.amount) || 0,
        visitor.paymentMethod?.toUpperCase() || '',
        visitor.paid ? 'Paid' : 'Unpaid',
      ])
    })
    const visitorWs = XLSX.utils.aoa_to_sheet(visitorRows)
    visitorWs['!cols'] = [
      { wch: 28 },
      { wch: 14 },
      { wch: 28 },
      { wch: 14 },
      { wch: 16 },
      { wch: 12 },
    ]
    XLSX.utils.book_append_sheet(wb, visitorWs, 'Visitors')

    // Sheet 3: Expenses
    const currentExpenses = getExpensesFor(year, month)
    const totalExpenses = currentExpenses.weeklyRent + currentExpenses.notablePrizes + currentExpenses.rosterSheet
    const netCollection = Math.max(0, summaryStats.grandTotalPayable - totalExpenses)

    const expensesRows = [
      ['Year', 'Month', 'Grand Total Payable', 'Weekly Rent', 'Notable Prizes', 'Roster Sheet', 'Total Expenses', 'Net Collection'],
      [
        year,
        MONTHS[month],
        summaryStats.grandTotalPayable,
        currentExpenses.weeklyRent,
        currentExpenses.notablePrizes,
        currentExpenses.rosterSheet,
        totalExpenses,
        netCollection,
      ],
    ]
    const expensesWs = XLSX.utils.aoa_to_sheet(expensesRows)
    expensesWs['!cols'] = [
      { wch: 8 },
      { wch: 12 },
      { wch: 20 },
      { wch: 14 },
      { wch: 16 },
      { wch: 14 },
      { wch: 16 },
      { wch: 16 },
    ]
    XLSX.utils.book_append_sheet(wb, expensesWs, 'Expenses')

    // Download master report
    XLSX.writeFile(wb, `BNI_Master_Report_${MONTHS[month]}_${year}.xlsx`)
  }

  const exportVisitorsExcel = () => {
    // Filter visitors for the selected month and year
    const filteredVisitors = visitors.filter((visitor) => {
      if (!visitor.friday) return false
      const d = new Date(visitor.friday)
      return d.getFullYear() === year && d.getMonth() === month
    })

    const wb = XLSX.utils.book_new()
    const rows = [
      ['Visitor Name', 'Friday', 'Referred By', 'Amount', 'Payment Method', 'Paid Status'],
      ...filteredVisitors.map((visitor) => [
        visitor.name,
        visitor.friday ? fmtDate(new Date(visitor.friday)) : '',
        visitor.referredBy || '',
        Number(visitor.amount) || 0,
        visitor.paymentMethod?.toUpperCase() || '',
        visitor.paid ? 'Paid' : 'Unpaid',
      ]),
    ]

    const ws = XLSX.utils.aoa_to_sheet(rows)
    ws['!cols'] = [
      { wch: 28 },
      { wch: 14 },
      { wch: 28 },
      { wch: 14 },
      { wch: 16 },
      { wch: 12 },
    ]
    XLSX.utils.book_append_sheet(wb, ws, 'Visitors')
    XLSX.writeFile(wb, `BNI_Visitors_${MONTHS[month]}_${year}.xlsx`)
  }

  const exportExpensesExcel = () => {
    const currentExpenses = getExpensesFor(year, month)
    const totalExpenses = currentExpenses.weeklyRent + currentExpenses.notablePrizes + currentExpenses.rosterSheet
    const netCollection = Math.max(0, summaryStats.grandTotalPayable - totalExpenses)

    const wb = XLSX.utils.book_new()
    const rows = [
      ['Year', 'Month', 'Grand Total Payable', 'Weekly Rent', 'Notable Prizes', 'Roster Sheet', 'Total Expenses', 'Net Collection'],
      [
        year,
        MONTHS[month],
        summaryStats.grandTotalPayable,
        currentExpenses.weeklyRent,
        currentExpenses.notablePrizes,
        currentExpenses.rosterSheet,
        totalExpenses,
        netCollection,
      ],
    ]

    const ws = XLSX.utils.aoa_to_sheet(rows)
    ws['!cols'] = [
      { wch: 8 },
      { wch: 12 },
      { wch: 20 },
      { wch: 14 },
      { wch: 16 },
      { wch: 14 },
      { wch: 16 },
      { wch: 16 },
    ]
    XLSX.utils.book_append_sheet(wb, ws, 'Expenses')
    XLSX.writeFile(wb, `BNI_Expenses_${MONTHS[month]}_${year}.xlsx`)
  }

  const addMember = () => {
    const name = newMemberName.trim()
    if (!name) {
      setMemberError('Enter a member name.')
      return
    }
    if (members.some((member) => member.toLowerCase() === name.toLowerCase())) {
      setMemberError('This member already exists.')
      return
    }
    setMembers((prev) => [...prev, name])
    setNewMemberName('')
    setMemberError('')
  }

  const startEditMember = (index) => {
    setEditingIndex(index)
    setEditingName(members[index])
    setMemberError('')
  }

  const saveEditedMember = () => {
    const name = editingName.trim()
    if (!name) {
      setMemberError('Enter a member name.')
      return
    }
    if (members.some((member, index) => index !== editingIndex && member.toLowerCase() === name.toLowerCase())) {
      setMemberError('This member already exists.')
      return
    }
    setMembers((prev) => prev.map((member, index) => (index === editingIndex ? name : member)))
    setEditingIndex(null)
    setEditingName('')
    setMemberError('')
  }

  const removeMember = (index) => {
    if (!window.confirm('Are you sure you want to remove this member?')) return

    setMembers((prev) => prev.filter((_, memberIndex) => memberIndex !== index))
    setAllData((prev) => {
      const next = {}
      Object.entries(prev).forEach(([dataYear, months]) => {
        next[dataYear] = {}
        Object.entries(months).forEach(([dataMonth, monthValue]) => {
          next[dataYear][dataMonth] = {
            ...monthValue,
            members: (monthValue.members ?? []).filter((_, memberIndex) => memberIndex !== index),
          }
        })
      })
      return next
    })
  }

  const addVisitor = () => {
    const name = visitorForm.name.trim()
    if (!name || !visitorForm.friday || !visitorForm.referredBy) {
      setVisitorError('Please enter visitor name, select Friday, and select referring member.')
      return
    }
    setVisitors((prev) => [
      ...prev,
      {
        id: makeId(),
        name,
        friday: visitorForm.friday,
        referredBy: visitorForm.referredBy,
        amount: Number(visitorForm.amount) || 300,
        paymentMethod: visitorForm.paymentMethod === 'upi' ? 'upi' : 'cash',
        paid: Boolean(visitorForm.paid),
      },
    ])
    setVisitorForm({ name: '', friday: visitorForm.friday, referredBy: '', amount: 300, paymentMethod: 'cash', paid: false })
    setVisitorError('')
  }

  const updateVisitor = (id, field, value) => {
    setVisitors((prev) => prev.map((visitor) => (
      visitor.id === id ? { ...visitor, [field]: value } : visitor
    )))
  }

  const removeVisitor = (id) => {
    if (!window.confirm('Are you sure you want to remove this visitor?')) return
    setVisitors((prev) => prev.filter((visitor) => visitor.id !== id))
  }

  if (!loggedIn) {
    return (
      <>
        <style>{css}</style>
        <div className="login-wrap">
          <div className="login-card">
            <div className="login-logo">BNI</div>
            <div className="login-sub">Amigos Chapter Admin Portal</div>
            <div className="field">
              <label htmlFor="login-user">Username</label>
              <input
                id="login-user"
                className="input"
                value={user}
                onChange={(event) => setUser(event.target.value)}
                onKeyDown={(event) => event.key === 'Enter' && handleLogin()}
                placeholder="Enter username"
                autoFocus
              />
            </div>
            <div className="field">
              <label htmlFor="login-pass">Password</label>
              <input
                id="login-pass"
                className="input"
                type="password"
                value={pass}
                onChange={(event) => setPass(event.target.value)}
                onKeyDown={(event) => event.key === 'Enter' && handleLogin()}
                placeholder="Enter password"
              />
            </div>
            <button className="primary-btn" style={{ width: '100%', marginTop: 8 }} onClick={handleLogin}>
              Login
            </button>
            {loginErr && <div className="login-error">{loginErr}</div>}
          </div>
        </div>
      </>
    )
  }

  return (
    <>
      <style>{css}</style>
      <div className="app-shell">
        <header className="topbar">
          <div className="brand">
            <span className="brand-mark">BNI</span>
            <span className="brand-sub">Amigos Chapter Attendance and Payment</span>
          </div>
          <nav className="nav" aria-label="Admin navigation">
            <button className={`nav-btn ${page === 'dashboard' ? 'active' : ''}`} onClick={() => setPage('dashboard')}>
              Dashboard
            </button>
            <button className={`nav-btn ${page === 'members' ? 'active' : ''}`} onClick={() => setPage('members')}>
              Members Management
            </button>
            <button className={`nav-btn ${page === 'visitors' ? 'active' : ''}`} onClick={() => setPage('visitors')}>
              Visitors
            </button>
            <button className={`nav-btn ${page === 'expenses' ? 'active' : ''}`} onClick={() => setPage('expenses')}>
              Expenses
            </button>
            <button
              className="ghost-btn"
              onClick={() => {
                setLoggedIn(false)
                setUser('')
                setPass('')
                setPage('dashboard')
              }}
            >
              Logout
            </button>
          </nav>
        </header>

        <main className="main">
          {page === 'dashboard' && (
            <DashboardPage
              year={year}
              month={month}
              years={years}
              members={members}
              monthData={monthData}
              contextMonthData={contextMonthData}
              visibleFridays={visibleFridays}
              fridayMode={fridayMode}
              paymentFilter={paymentFilter}
              paymentStatus={paymentStatus}
              summaryStats={summaryStats}
              displayedMemberIndexes={displayedMemberIndexes}
              rowTotals={rowTotals}
              rowPaidTotals={rowPaidTotals}
              visibleGrandTotal={visibleGrandTotal}
              visiblePaidTotal={visiblePaidTotal}
              getFridayEntry={getFridayEntry}
              setYear={setYear}
              setMonth={setMonth}
              setFridayMode={setFridayMode}
              setPaymentFilter={setPaymentFilter}
              setPaymentStatus={setPaymentStatus}
              updatePayMode={updatePayMode}
              updateFridayField={updateFridayField}
              updateRent={updateRent}
              exportExcel={exportExcel}
            />
          )}

          {page === 'members' && (
            <MembersPage
              members={members}
              newMemberName={newMemberName}
              memberError={memberError}
              editingIndex={editingIndex}
              editingName={editingName}
              setNewMemberName={setNewMemberName}
              setEditingName={setEditingName}
              addMember={addMember}
              startEditMember={startEditMember}
              saveEditedMember={saveEditedMember}
              cancelEdit={() => {
                setEditingIndex(null)
                setEditingName('')
                setMemberError('')
              }}
              removeMember={removeMember}
            />
          )}

          {page === 'visitors' && (
            <VisitorsPage
              year={year}
              month={month}
              members={members}
              visitors={visitors}
              visitorForm={visitorForm}
              visitorError={visitorError}
              setVisitorForm={setVisitorForm}
              addVisitor={addVisitor}
              updateVisitor={updateVisitor}
              removeVisitor={removeVisitor}
              exportVisitorsExcel={exportVisitorsExcel}
            />
          )}

          {page === 'expenses' && (
            <ExpensesPage
              year={year}
              month={month}
              years={years}
              setYear={setYear}
              setMonth={setMonth}
              grandTotalPayable={summaryStats.grandTotalPayable}
              getExpensesFor={getExpensesFor}
              setExpensesFor={setExpensesFor}
              exportExpensesExcel={exportExpensesExcel}
            />
          )}
        </main>
      </div>
    </>
  )
}

function DashboardPage(props) {
  const {
    year,
    month,
    years,
    members,
    monthData,
    contextMonthData,
    visibleFridays,
    fridayMode,
    paymentFilter,
    paymentStatus,
    summaryStats,
    displayedMemberIndexes,
    rowTotals,
    rowPaidTotals,
    visibleGrandTotal,
    visiblePaidTotal,
    getFridayEntry,
    setYear,
    setMonth,
    setFridayMode,
    setPaymentFilter,
    setPaymentStatus,
    updatePayMode,
    updateFridayField,
    updateRent,
    exportExcel,
  } = props

  return (
    <>
      <div className="page-head">
        <div>
          <h1 className="page-title">Attendance Dashboard</h1>
          <p className="page-copy">Track member attendance, payment mode, paid status, and venue rent.</p>
        </div>
      </div>

      <section className="toolbar" aria-label="Dashboard controls">
        <div className="control compact">
          <span className="form-label">Year</span>
          <select className="select" value={year} onChange={(event) => setYear(Number(event.target.value))}>
            {years.map((item) => <option key={item} value={item}>{item}</option>)}
          </select>
        </div>
        <div className="control">
          <span className="form-label">Month</span>
          <select className="select" value={month} onChange={(event) => setMonth(Number(event.target.value))}>
            {MONTHS.map((name, index) => <option key={name} value={index}>{name}</option>)}
          </select>
        </div>
        <div className="control">
          <span className="form-label">Friday View</span>
          <div className="segmented" role="group" aria-label="Friday view">
            <button
              className={`segment ${fridayMode === 'current' ? 'active' : ''}`}
              onClick={() => setFridayMode('current')}
            >
              Current / Upcoming Friday
            </button>
            <button
              className={`segment ${fridayMode === 'all' ? 'active' : ''}`}
              onClick={() => setFridayMode('all')}
            >
              All Fridays
            </button>
          </div>
        </div>
        <div className="control compact">
          <span className="form-label">Payment Type</span>
          <select className="select" value={paymentFilter} onChange={(event) => setPaymentFilter(event.target.value)}>
            <option value="all">All</option>
            <option value="weekly">Weekly</option>
            <option value="monthly">Monthly</option>
          </select>
        </div>
        <div className="control compact">
          <span className="form-label">Payment Status</span>
          <select className="select" value={paymentStatus} onChange={(event) => setPaymentStatus(event.target.value)}>
            <option value="all">All</option>
            <option value="paid">Paid</option>
            <option value="unpaid">Unpaid</option>
          </select>
        </div>
        <div className="quick-actions" aria-label="Quick payment filters">
          <button className="muted-btn" onClick={() => setPaymentStatus('paid')}>View Paid Members</button>
          <button className="muted-btn" onClick={() => setPaymentStatus('unpaid')}>View Unpaid Members</button>
        </div>
        <button className="primary-btn" onClick={exportExcel}>Download Excel</button>
        <span className="status-note">
          Showing {visibleFridays.map((friday) => fmtDate(friday.date)).join(', ')}
        </span>
      </section>

      <section className="summary-grid" aria-label="Payment summary">
        <div className="summary-card">
          <div className="summary-label">Total Members Paid</div>
          <div className="summary-value success">{summaryStats.totalMembersPaid}</div>
        </div>
        <div className="summary-card">
          <div className="summary-label">Members Collection</div>
          <div className="summary-value success">₹{summaryStats.memberCollection.toLocaleString('en-IN')}</div>
        </div>
        <div className="summary-card">
          <div className="summary-label">Net Collection</div>
          <div className="summary-value success">₹{summaryStats.netCollection.toLocaleString('en-IN')}</div>
        </div>
        
        
        <div className="summary-card">
          <div className="summary-label">Outstanding Amount(Remaining Amount)</div>
          <div className="summary-value warning">₹{summaryStats.outstandingAmount.toLocaleString('en-IN')}</div>
        </div>
        <div className="summary-card">
          <div className="summary-label">Total Members Unpaid</div>
          <div className="summary-value warning">{summaryStats.totalMembersUnpaid}</div>
        </div>
        <div className="summary-card">
          <div className="summary-label">Visitors Collection</div>
          <div className="summary-value success">₹{summaryStats.visitorCollection.toLocaleString('en-IN')}</div>
        </div>
        
      </section>

      <section className="panel rent-panel">
        <div>
          <h2 className="panel-title">Monthly Meeting Venue Rent</h2>
          <p className="panel-copy">Amount payable for meetings held in {MONTHS[month]} {year}.</p>
        </div>
        <div className="rent-fields">
          <label className="form-label" htmlFor="rent-amount">Amount</label>
          <input
            id="rent-amount"
            className="input amount-input"
            type="number"
            value={monthData.rent.amount}
            onChange={(event) => updateRent('amount', Number(event.target.value))}
          />
          <label className="form-label" htmlFor="rent-status">Status</label>
          <select
            id="rent-status"
            className="select"
            value={monthData.rent.status}
            onChange={(event) => updateRent('status', event.target.value)}
          >
            <option value="not_paid">Not Paid</option>
            <option value="paid">Paid</option>
          </select>
          <span className={`badge ${monthData.rent.status === 'paid' ? 'ok' : 'warn'}`}>
            {monthData.rent.status === 'paid' ? 'Paid' : 'Pending'} - Rs. {Number(monthData.rent.amount).toLocaleString('en-IN')}
          </span>
        </div>
      </section>

      <section className="panel legend" aria-label="Payment rules">
        <strong>Rules:</strong>
        <span className="legend-item"><span className="legend-dot" style={{ background: '#166534' }} /> Weekly: Present ₹800, Monthly: Present ₹0 (except Friday 1: ₹3000/₹4300)</span>
        <span className="legend-item"><span className="legend-dot" style={{ background: '#991b1b' }} /> Weekly Absent: ₹1200, Monthly Absent: ₹400</span>
        <span className="legend-item"><span className="legend-dot" style={{ background: '#92400e' }} /> Weekly Substitute: ₹1000, Monthly Substitute: ₹200</span>
      </section>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th rowSpan={2}>#</th>
              <th className="name-head" rowSpan={2}>Member Name</th>
              <th rowSpan={2}>Pay Mode</th>
              {visibleFridays.map((friday, index) => (
                <th key={`${friday.year}-${friday.month}-${friday.index}`} colSpan={3}>
                  Friday {index + 1}<br />
                  <span style={{ fontWeight: 600, fontSize: '0.74rem' }}>{fmtDate(friday.date)}</span>
                </th>
              ))}
              <th rowSpan={2}>Total Payable</th>
              <th rowSpan={2}>Due</th>
              <th rowSpan={2}>Total Paid</th>
            </tr>
            <tr>
              {visibleFridays.map((friday) => (
                <Fragment key={`sub-${friday.year}-${friday.month}-${friday.index}`}>
                  <th>Attendance</th>
                  <th>Amount</th>
                  <th>Payment</th>
                </Fragment>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayedMemberIndexes.length === 0 && (
              <tr>
                <td colSpan={6 + visibleFridays.length * 3} className="empty-state">
                  No members match the selected filters.
                </td>
              </tr>
            )}

            {displayedMemberIndexes.map((memberIndex, visibleIndex) => {
              const memberRow = contextMonthData.members[memberIndex] ?? initMemberRow(visibleFridays.length)
              return (
                <tr key={`${members[memberIndex]}-${memberIndex}`}>
                  <td className="sno-cell">{visibleIndex + 1}</td>
                  <td className="name-cell">{members[memberIndex]}</td>
                  <td>
                    <div className="paymode-toggle" role="group" aria-label={`Payment mode for ${members[memberIndex]}`}>
                      <button
                        className={`paymode-btn ${memberRow.payMode === 'weekly' ? 'active' : ''}`}
                        onClick={() => updatePayMode(memberIndex, 'weekly')}
                      >
                        Weekly
                      </button>
                      <button
                        className={`paymode-btn ${memberRow.payMode === 'monthly' ? 'active' : ''}`}
                        onClick={() => updatePayMode(memberIndex, 'monthly')}
                      >
                        Monthly
                      </button>
                    </div>
                  </td>
                  {visibleFridays.map((friday) => {
                    const entry = getFridayEntry(memberIndex, friday)
                    const config = STATUS_CONFIG[entry.status] ?? STATUS_CONFIG.present
                    return (
                      <Fragment key={`${memberIndex}-${friday.year}-${friday.month}-${friday.index}`}>
                        <td>
                          <select
                            className="mini-select"
                            value={entry.status}
                            onChange={(event) => updateFridayField(memberIndex, friday, 'status', event.target.value)}
                            style={{ color: config.color, background: config.bg }}
                          >
                            {Object.entries(STATUS_CONFIG).map(([key, value]) => (
                              <option key={key} value={key}>{value.label}</option>
                            ))}
                          </select>
                        </td>
                        <td>
                          <input
                            className="mini-input"
                            type="number"
                            value={entry.amount}
                            onChange={(event) => updateFridayField(memberIndex, friday, 'amount', Number(event.target.value))}
                            style={{ color: config.color }}
                          />
                        </td>
                        <td>
                          <div className="payment-cell">
                            <select
                              className="mini-select"
                              value={entry.payMethod}
                              onChange={(event) => updateFridayField(memberIndex, friday, 'payMethod', event.target.value)}
                            >
                              <option value="cash">Cash</option>
                              <option value="upi">UPI</option>
                            </select>
                            <label className="paid-check">
                              <input
                                type="checkbox"
                                checked={entry.paid}
                                onChange={(event) => updateFridayField(memberIndex, friday, 'paid', event.target.checked)}
                              />
                              Paid
                            </label>
                          </div>
                        </td>
                      </Fragment>
                    )
                  })}
                  <td className="total-cell">Rs. {rowTotals[memberIndex].toLocaleString('en-IN')}</td>
                  <td className="total-cell">Rs. {Math.max(0, rowTotals[memberIndex] - rowPaidTotals[memberIndex]).toLocaleString('en-IN')}</td>
                  <td className="paid-total-cell">Rs. {rowPaidTotals[memberIndex].toLocaleString('en-IN')}</td>
                </tr>
              )
            })}

            <tr className="summary-row">
              <td></td>
              <td>Grand Total</td>
              <td></td>
              {visibleFridays.flatMap((friday) => [
                <td key={`s-${friday.year}-${friday.month}-${friday.index}`}></td>,
                <td key={`a-${friday.year}-${friday.month}-${friday.index}`}></td>,
                <td key={`p-${friday.year}-${friday.month}-${friday.index}`}></td>,
              ])}
              <td>Rs. {visibleGrandTotal.toLocaleString('en-IN')}</td>
              <td>Rs. {Math.max(0, visibleGrandTotal - visiblePaidTotal).toLocaleString('en-IN')}</td>
              <td>Rs. {visiblePaidTotal.toLocaleString('en-IN')}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <section className="qr-section">
        <h2 className="qr-title">Payment - PhonePe / UPI</h2>
        <p className="qr-sub">Scan this QR code using PhonePe or any UPI app to make a payment.</p>
        <div className="qr-img-wrap">
          <img src={QR_BASE64} alt="PhonePe QR Code" className="qr-img" />
        </div>
        <div className="qr-name">Mr. Vobbilisetty Chandra Krishna</div>
      </section>
    </>
  )
}

function MembersPage(props) {
  const {
    members,
    newMemberName,
    memberError,
    editingIndex,
    editingName,
    setNewMemberName,
    setEditingName,
    addMember,
    startEditMember,
    saveEditedMember,
    cancelEdit,
    removeMember,
  } = props

  return (
    <>
      <div className="page-head">
        <div>
          <h1 className="page-title">Members Management</h1>
          <p className="page-copy">Add, rename, and remove active chapter members.</p>
        </div>
      </div>

      <section className="panel">
        <div className="form-grid">
          <div className="field" style={{ marginBottom: 0 }}>
            <label htmlFor="new-member">Member Name</label>
            <input
              id="new-member"
              className="input"
              value={newMemberName}
              onChange={(event) => setNewMemberName(event.target.value)}
              onKeyDown={(event) => event.key === 'Enter' && addMember()}
              placeholder="Enter member name"
            />
          </div>
          <button className="primary-btn" onClick={addMember}>Save Member</button>
        </div>
        {memberError && <div className="form-error">{memberError}</div>}
      </section>

      <section className="management-list" aria-label="Members list">
        {members.map((member, index) => (
          <div className="member-row" key={`${member}-${index}`}>
            <div className="avatar">{initials(member)}</div>
            {editingIndex === index ? (
              <input
                className="input"
                value={editingName}
                onChange={(event) => setEditingName(event.target.value)}
                onKeyDown={(event) => event.key === 'Enter' && saveEditedMember()}
                autoFocus
              />
            ) : (
              <strong>{member}</strong>
            )}
            <div className="row-actions">
              {editingIndex === index ? (
                <>
                  <button className="primary-btn" onClick={saveEditedMember}>Save</button>
                  <button className="muted-btn" onClick={cancelEdit}>Cancel</button>
                </>
              ) : (
                <>
                  <button className="muted-btn" onClick={() => startEditMember(index)}>Edit</button>
                  <button className="danger-btn" onClick={() => removeMember(index)}>Remove</button>
                </>
              )}
            </div>
          </div>
        ))}
      </section>
    </>
  )
}

function VisitorsPage(props) {
  const {
    year,
    month,
    members,
    visitors,
    visitorForm,
    visitorError,
    setVisitorForm,
    addVisitor,
    updateVisitor,
    removeVisitor,
    exportVisitorsExcel,
  } = props

  const total = visitors.reduce((sum, visitor) => sum + (Number(visitor.amount) || 0), 0)
  const paidTotal = visitors.reduce((sum, visitor) => sum + (visitor.paid ? Number(visitor.amount) || 0 : 0), 0)

  // Get all Fridays for current and next month
  const allFridays = useMemo(() => {
    const fridays = []
    // Add current month
    getFridays(year, month).forEach((d) => fridays.push(d))
    // Add next month
    const nextMonth = month === 11 ? 0 : month + 1
    const nextYear = month === 11 ? year + 1 : year
    getFridays(nextYear, nextYear).forEach((d) => fridays.push(d))
    return fridays
  }, [year, month])

  return (
    <>
      <div className="page-head">
        <div>
          <h1 className="page-title">Visitors</h1>
          <p className="page-copy">Track visitor payments separately from chapter member records.</p>
        </div>
        <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
          <div className="badge ok">Paid Rs. {paidTotal.toLocaleString('en-IN')} of Rs. {total.toLocaleString('en-IN')}</div>
          <button className="primary-btn" onClick={exportVisitorsExcel}>Download Visitors Excel</button>
        </div>
      </div>

      <section className="panel">
        <div className="form-grid visitor-form-grid" style={{ gridTemplateColumns: 'minmax(220px, 1fr) 150px minmax(200px, 1fr) 120px 130px 90px auto' }}>
          <div className="field" style={{ marginBottom: 0 }}>
            <label htmlFor="visitor-name">Visitor Name</label>
            <input
              id="visitor-name"
              className="input"
              value={visitorForm.name}
              onChange={(event) => setVisitorForm((prev) => ({ ...prev, name: event.target.value }))}
              placeholder="Enter visitor name"
            />
          </div>
          <div className="field" style={{ marginBottom: 0 }}>
            <label htmlFor="visitor-friday">Friday</label>
            <select
              id="visitor-friday"
              className="select"
              value={visitorForm.friday}
              onChange={(event) => setVisitorForm((prev) => ({ ...prev, friday: event.target.value }))}
            >
              <option value="">Select Friday</option>
              {allFridays.map((d) => (
                <option key={d.toISOString()} value={d.toISOString()}>
                  {fmtDate(d)}
                </option>
              ))}
            </select>
          </div>
          <div className="field" style={{ marginBottom: 0 }}>
            <label htmlFor="visitor-referred">Referred By</label>
            <select
              id="visitor-referred"
              className="select"
              value={visitorForm.referredBy}
              onChange={(event) => setVisitorForm((prev) => ({ ...prev, referredBy: event.target.value }))}
            >
              <option value="">Select Member</option>
              {members.map((member) => (
                <option key={member} value={member}>{member}</option>
              ))}
            </select>
          </div>
          <div className="field" style={{ marginBottom: 0 }}>
            <label htmlFor="visitor-amount">Amount</label>
            <input
              id="visitor-amount"
              className="input"
              type="number"
              value={visitorForm.amount}
              onChange={(event) => setVisitorForm((prev) => ({ ...prev, amount: Number(event.target.value) }))}
            />
          </div>
          <div className="field" style={{ marginBottom: 0 }}>
            <label htmlFor="visitor-payment">Payment</label>
            <select
              id="visitor-payment"
              className="select"
              value={visitorForm.paymentMethod}
              onChange={(event) => setVisitorForm((prev) => ({ ...prev, paymentMethod: event.target.value }))}
            >
              <option value="cash">Cash</option>
              <option value="upi">UPI</option>
            </select>
          </div>
          <label className="paid-check" style={{ alignSelf: 'center', marginTop: 21 }}>
            <input
              type="checkbox"
              checked={visitorForm.paid}
              onChange={(event) => setVisitorForm((prev) => ({ ...prev, paid: event.target.checked }))}
            />
            Paid
          </label>
          <button className="primary-btn" onClick={addVisitor}>Add Visitor</button>
        </div>
        {visitorError && <div className="form-error">{visitorError}</div>}
      </section>

      <section className="management-list" aria-label="Visitors list">
        {visitors.length === 0 && (
          <div className="panel empty-state">No visitors have been added.</div>
        )}
        {visitors.map((visitor) => (
          <div className="visitor-row" key={visitor.id} style={{ gridTemplateColumns: 'minmax(180px, 1fr) 150px minmax(200px, 1fr) 120px 130px 90px auto' }}>
            <input
              className="input"
              value={visitor.name}
              onChange={(event) => updateVisitor(visitor.id, 'name', event.target.value)}
            />
            <select
              className="select"
              value={visitor.friday || ''}
              onChange={(event) => updateVisitor(visitor.id, 'friday', event.target.value)}
            >
              <option value="">Select Friday</option>
              {allFridays.map((d) => (
                <option key={d.toISOString()} value={d.toISOString()}>
                  {fmtDate(d)}
                </option>
              ))}
            </select>
            <select
              className="select"
              value={visitor.referredBy || ''}
              onChange={(event) => updateVisitor(visitor.id, 'referredBy', event.target.value)}
            >
              <option value="">Select Member</option>
              {members.map((member) => (
                <option key={member} value={member}>{member}</option>
              ))}
            </select>
            <input
              className="input"
              type="number"
              value={visitor.amount}
              onChange={(event) => updateVisitor(visitor.id, 'amount', Number(event.target.value))}
            />
            <select
              className="select"
              value={visitor.paymentMethod}
              onChange={(event) => updateVisitor(visitor.id, 'paymentMethod', event.target.value)}
            >
              <option value="cash">Cash</option>
              <option value="upi">UPI</option>
            </select>
            <label className="paid-check">
              <input
                type="checkbox"
                checked={visitor.paid}
                onChange={(event) => updateVisitor(visitor.id, 'paid', event.target.checked)}
              />
              Paid
            </label>
            <div className="row-actions">
              <button className="danger-btn" onClick={() => removeVisitor(visitor.id)}>Remove</button>
            </div>
          </div>
        ))}
      </section>
    </>
  )
}

function ExpensesPage(props) {
  const {
    year,
    month,
    years,
    setYear,
    setMonth,
    grandTotalPayable,
    getExpensesFor,
    setExpensesFor,
    exportExpensesExcel,
  } = props

  const [weeklyRent, setWeeklyRent] = useState(0)
  const [notablePrizes, setNotablePrizes] = useState(0)
  const [rosterSheet, setRosterSheet] = useState(0)
  const [showBanner, setShowBanner] = useState(false)
  const isApplyingRef = useRef(false)
  const timeoutRef = useRef(null)

  // Cleanup timeout on unmount
  useEffect(() => {
    return () => {
      if (timeoutRef.current) {
        clearTimeout(timeoutRef.current)
      }
    }
  }, [])

  useEffect(() => {
    const current = getExpensesFor(year, month)
    setWeeklyRent(current.weeklyRent)
    setNotablePrizes(current.notablePrizes)
    setRosterSheet(current.rosterSheet)

    if (!isApplyingRef.current) {
      setShowBanner(false)
    } else {
      isApplyingRef.current = false
    }
  }, [year, month, getExpensesFor])

  const totalExpenses = weeklyRent + notablePrizes + rosterSheet
  const netCollection = Math.max(0, grandTotalPayable - totalExpenses)

  const applyExpenses = () => {
    isApplyingRef.current = true
    setExpensesFor(year, month, { weeklyRent, notablePrizes, rosterSheet })
    setShowBanner(true)

    if (timeoutRef.current) {
      clearTimeout(timeoutRef.current)
    }

    timeoutRef.current = setTimeout(() => {
      setShowBanner(false)
    }, 3000)
  }

  return (
    <>
      <div className="page-head">
        <div>
          <h1 className="page-title">Expenses</h1>
          <p className="page-copy">Track and manage monthly expenses.</p>
        </div>
      </div>

      <section className="toolbar" aria-label="Expense controls">
        <div className="control compact">
          <span className="form-label">Year</span>
          <select className="select" value={year} onChange={(event) => setYear(Number(event.target.value))}>
            {years.map((item) => <option key={item} value={item}>{item}</option>)}
          </select>
        </div>
        <div className="control">
          <span className="form-label">Month</span>
          <select className="select" value={month} onChange={(event) => setMonth(Number(event.target.value))}>
            {MONTHS.map((name, index) => <option key={name} value={index}>{name}</option>)}
          </select>
        </div>
      </section>

      {/* Hero Card - Net Collection */}
      <section style={{ marginBottom: '16px' }}>
        <div style={{
          background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
          borderRadius: '12px',
          padding: '32px',
          textAlign: 'center',
          color: 'white',
          boxShadow: '0 10px 25px rgba(16, 185, 129, 0.25)'
        }}>
          <div style={{ fontSize: '0.85rem', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: '8px', fontWeight: 'bold', opacity: 0.9 }}>
            Net Collection
          </div>
          <div style={{ fontSize: '3rem', fontWeight: '800', lineHeight: '1' }}>
            ₹{netCollection.toLocaleString('en-IN')}
          </div>
        </div>
      </section>

      {/* Financial Summary Cards */}
      <section className="summary-grid" aria-label="Financial summary">
        <div className="summary-card">
          <div className="summary-label">Grand Total Payable</div>
          <div className="summary-value">₹{grandTotalPayable.toLocaleString('en-IN')}</div>
        </div>
        <div className="summary-card">
          <div className="summary-label">Total Expenses</div>
          <div className="summary-value warning">₹{totalExpenses.toLocaleString('en-IN')}</div>
        </div>
      </section>

      {/* Expense Breakdown Card */}
      <section className="panel">
        <h2 className="panel-title" style={{ marginBottom: '20px' }}>Expense Breakdown</h2>
        
        <div style={{ display: 'grid', gap: '16px', marginBottom: '24px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <label style={{ minWidth: '120px', fontWeight: '700', color: '#627083' }}>Weekly Rent</label>
            <span style={{ fontSize: '1.1rem', fontWeight: '700', marginRight: '8px' }}>₹</span>
            <input
              id="weekly-rent"
              className="input"
              type="number"
              value={weeklyRent}
              onChange={(event) => setWeeklyRent(Number(event.target.value))}
              style={{ maxWidth: '200px' }}
            />
          </div>
          
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <label style={{ minWidth: '120px', fontWeight: '700', color: '#627083' }}>Notable Prizes</label>
            <span style={{ fontSize: '1.1rem', fontWeight: '700', marginRight: '8px' }}>₹</span>
            <input
              id="notable-prizes"
              className="input"
              type="number"
              value={notablePrizes}
              onChange={(event) => setNotablePrizes(Number(event.target.value))}
              style={{ maxWidth: '200px' }}
            />
          </div>
          
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <label style={{ minWidth: '120px', fontWeight: '700', color: '#627083' }}>Roster Sheet</label>
            <span style={{ fontSize: '1.1rem', fontWeight: '700', marginRight: '8px' }}>₹</span>
            <input
              id="roster-sheet"
              className="input"
              type="number"
              value={rosterSheet}
              onChange={(event) => setRosterSheet(Number(event.target.value))}
              style={{ maxWidth: '200px' }}
            />
          </div>
        </div>

        {/* Calculation Preview */}
        <div style={{
          background: '#f9fbfd',
          borderRadius: '8px',
          padding: '20px',
          marginBottom: '20px',
          border: '1px solid #e4eaf2'
        }}>
          <div style={{ display: 'grid', gap: '8px', maxWidth: '400px', margin: '0 auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '700' }}>
              <span>Grand Total Payable</span>
              <span>₹{grandTotalPayable.toLocaleString('en-IN')}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '700', color: '#991b1b' }}>
              <span>Total Expenses</span>
              <span>₹{totalExpenses.toLocaleString('en-IN')}</span>
            </div>
            <div style={{ borderTop: '2px solid #627083', marginTop: '8px', paddingTop: '8px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: '800', fontSize: '1.1rem' }}>
                <span>Net Collection</span>
                <span style={{ color: '#10b981' }}>₹{netCollection.toLocaleString('en-IN')}</span>
              </div>
            </div>
          </div>
        </div>

        {/* Success Banner */}
        {showBanner && (
          <div style={{
            background: '#10b981',
            color: 'white',
            padding: '12px 24px',
            borderRadius: '8px',
            marginBottom: '20px',
            textAlign: 'center',
            fontWeight: '700',
            fontSize: '1.1rem'
          }}>
            Expenses Applied Successfully!
          </div>
        )}

        {/* Buttons */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ flex: 1 }}></div>
          <button className="primary-btn" onClick={applyExpenses} style={{ padding: '12px 40px', fontSize: '1.1rem' }}>
            Apply Expenses
          </button>
          <div style={{ flex: 1, display: 'flex', justifyContent: 'flex-end' }}>
            <button className="primary-btn" onClick={exportExpensesExcel} style={{ padding: '12px 40px', fontSize: '1.1rem' }}>
              Download Expenses Excel
            </button>
          </div>
        </div>
      </section>
    </>
  )
}
