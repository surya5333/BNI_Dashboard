import { useState, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import QR_BASE64 from './qrData.js'

// ─── CONSTANTS ────────────────────────────────────────────────────────────────

const MEMBERS = [
  "Dr. B. Anusha Balaram","Chandu Krishna","K. V. T. Ramesh","Rama Devi",
  "K. V. T. Chandra Mouli","Tarak","Naveen","Nageswara Rao","Rajesh Kumar V",
  "Manikanta Teja K","Ajay Kumar P","Rajesh V","Naveen Kumar K",
  "Dhilleswara Rao D","Santosh Kumar G","Venkateswarlu T","Ravikumar G",
  "Kalyani B","Viswa Sai Dhanush K","Balaram","V. Chandra Krishna",
  "Sateesh A","Kalyan Bogi","Nagraju G","Gopal","Dr. John","Ramesh T",
  "S. Udaya Bhaskar Rao","Nageswararao K","Ajay Kumar","Kalkinadh Gudla",
  "Santosh Kumar P","Rajasekhar V","Santosh Reddi","Uday K",
  "Tharak Pydisetty","Bhavani","Manikanta","V. Santosh Kumar","Anusha"
]

const MONTHS = ["January","February","March","April","May","June",
  "July","August","September","October","November","December"]

const STATUS_CONFIG = {
  present:    { label: '✅ Present',    color: '#16a34a', bg: '#dcfce7' },
  absent:     { label: '❌ Absent',     color: '#dc2626', bg: '#fee2e2' },
  substitute: { label: '🔄 Substitute', color: '#d97706', bg: '#fef3c7' },
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────

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

function fmtDate(d) {
  return `${d.getDate()} ${MONTHS[d.getMonth()].slice(0,3)}`
}

function calcAmount(status, payMode, totalFridays) {
  if (status === 'absent') return 1200
  if (status === 'substitute') return 1000
  if (payMode === 'weekly') return 800
  return totalFridays >= 5 ? 4300 : 3000
}

function initMemberRow(fridayCount) {
  return {
    payMode: 'weekly',
    fridays: Array.from({ length: fridayCount }, () => ({
      status: 'present',
      amount: 800,
      payMethod: 'cash',
    })),
  }
}

function initMonthData(year, month) {
  const count = getFridays(year, month).length
  return {
    members: MEMBERS.map(() => initMemberRow(count)),
    rent: { amount: 13000, status: 'not_paid' },
  }
}

// ─── GLOBAL STYLES ─────────────────────────────────────────────────────────────

const G = {
  red:    '#c8102e',
  navy:   '#1a2233',
  blue:   '#2d4a7a',
  gold:   '#f0b429',
  green:  '#16a34a',
  lightBg:'#f0f4f8',
  white:  '#ffffff',
  border: '#dde3ee',
  text:   '#1a2233',
  muted:  '#64748b',
}

const css = `
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', system-ui, sans-serif; background: ${G.lightBg}; color: ${G.text}; }
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: #e2e8f0; }
  ::-webkit-scrollbar-thumb { background: ${G.blue}; border-radius: 4px; }

  .login-wrap {
    min-height: 100vh; display: flex; align-items: center; justify-content: center;
    background: linear-gradient(135deg, ${G.navy} 0%, #1e3a6e 60%, #2d4a7a 100%);
    position: relative; overflow: hidden;
  }
  .login-wrap::before {
    content: 'BNI'; position: absolute; font-family: 'Syne', sans-serif;
    font-size: 40vw; font-weight: 900; color: rgba(255,255,255,0.025);
    user-select: none; top: 50%; left: 50%; transform: translate(-50%,-50%);
    letter-spacing: -2vw;
  }
  .login-card {
    background: #fff; border-radius: 20px; padding: 52px 44px 44px;
    width: 400px; max-width: 94vw;
    box-shadow: 0 24px 80px rgba(0,0,0,0.35); position: relative; z-index: 1;
  }
  .login-logo {
    font-family: 'Syne', sans-serif; font-size: 3rem; font-weight: 900;
    color: ${G.red}; text-align: center; letter-spacing: 4px; line-height: 1;
  }
  .login-sub {
    text-align: center; color: ${G.blue}; font-size: 0.88rem; font-weight: 600;
    letter-spacing: 0.5px; margin-top: 6px; margin-bottom: 36px;
    text-transform: uppercase;
  }
  .login-field { margin-bottom: 14px; }
  .login-field label { display: block; font-size: 0.78rem; font-weight: 700;
    color: ${G.muted}; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px; }
  .login-field input {
    width: 100%; padding: 12px 16px; border: 1.5px solid ${G.border};
    border-radius: 10px; font-size: 1rem; outline: none;
    transition: border-color 0.2s, box-shadow 0.2s; font-family: inherit;
  }
  .login-field input:focus { border-color: ${G.blue}; box-shadow: 0 0 0 3px rgba(45,74,122,0.12); }
  .login-btn {
    width: 100%; padding: 14px; margin-top: 8px;
    background: linear-gradient(135deg, ${G.red}, #a00d25);
    color: #fff; border: none; border-radius: 10px; font-size: 1rem;
    font-weight: 800; letter-spacing: 2px; cursor: pointer;
    transition: transform 0.15s, box-shadow 0.15s; font-family: inherit;
  }
  .login-btn:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(200,16,46,0.4); }
  .login-btn:active { transform: translateY(0); }
  .login-err { color: ${G.red}; font-size: 0.85rem; text-align: center; margin-top: 12px; font-weight: 600; }

  /* TOPBAR */
  .topbar {
    position: sticky; top: 0; z-index: 200;
    background: linear-gradient(90deg, ${G.navy} 0%, ${G.blue} 100%);
    height: 60px; display: flex; align-items: center; justify-content: space-between;
    padding: 0 28px; box-shadow: 0 2px 16px rgba(0,0,0,0.22);
  }
  .topbar-left { display: flex; align-items: center; gap: 14px; }
  .topbar-bni { font-family: 'Syne', sans-serif; font-size: 1.5rem; font-weight: 900; color: ${G.gold}; letter-spacing: 3px; }
  .topbar-chapter { font-size: 0.85rem; color: #93b4d8; font-weight: 500;
    border-left: 1.5px solid #3d5a8a; padding-left: 14px; }
  .topbar-right { display: flex; align-items: center; gap: 12px; }
  .badge-admin {
    background: ${G.red}; color: #fff; padding: 5px 14px;
    border-radius: 20px; font-size: 0.78rem; font-weight: 800; letter-spacing: 0.5px;
  }
  .btn-logout {
    border: 1.5px solid ${G.gold}; color: ${G.gold}; background: transparent;
    padding: 5px 14px; border-radius: 20px; font-size: 0.78rem; font-weight: 700;
    cursor: pointer; transition: all 0.2s; font-family: inherit;
  }
  .btn-logout:hover { background: ${G.gold}; color: ${G.navy}; }

  /* MAIN */
  .main { padding: 24px 20px 80px; max-width: 1700px; margin: 0 auto; }

  /* CONTROLS */
  .controls-row {
    display: flex; align-items: center; gap: 14px; flex-wrap: wrap; margin-bottom: 18px;
  }
  .ctrl-group { display: flex; align-items: center; gap: 8px; }
  .ctrl-label { font-size: 0.78rem; font-weight: 800; color: ${G.blue};
    text-transform: uppercase; letter-spacing: 0.6px; }
  .ctrl-select {
    padding: 9px 14px; border: 1.5px solid ${G.border}; border-radius: 9px;
    font-size: 0.9rem; background: #fff; outline: none; font-family: inherit;
    cursor: pointer; transition: border-color 0.2s;
  }
  .ctrl-select:focus { border-color: ${G.blue}; }
  .btn {
    padding: 9px 20px; border-radius: 9px; font-size: 0.85rem; font-weight: 800;
    cursor: pointer; border: none; letter-spacing: 0.5px; font-family: inherit;
    transition: all 0.2s;
  }
  .btn-dl { background: ${G.green}; color: #fff; }
  .btn-dl:hover { background: #15803d; box-shadow: 0 4px 14px rgba(22,163,74,0.35); }

  /* RENT CARD */
  .rent-card {
    background: #fff; border-radius: 14px; padding: 18px 24px;
    margin-bottom: 18px; display: flex; align-items: center; gap: 20px; flex-wrap: wrap;
    border-left: 5px solid ${G.gold}; box-shadow: 0 2px 10px rgba(0,0,0,0.07);
  }
  .rent-title { font-size: 0.95rem; font-weight: 800; color: ${G.navy}; }
  .rent-sub { font-size: 0.78rem; color: ${G.muted}; margin-top: 2px; }
  .rent-amount-row { display: flex; align-items: center; gap: 6px; }
  .rent-rupee { font-weight: 900; color: ${G.green}; font-size: 1.1rem; }
  .rent-input {
    width: 110px; padding: 8px 10px; border: 1.5px solid ${G.border}; border-radius: 8px;
    font-size: 1rem; font-weight: 700; color: ${G.green}; text-align: center;
    outline: none; font-family: inherit;
  }
  .rent-input:focus { border-color: ${G.green}; }
  .rent-sel {
    padding: 8px 12px; border: 1.5px solid ${G.border}; border-radius: 8px;
    font-size: 0.9rem; font-weight: 600; font-family: inherit; cursor: pointer;
    outline: none; background: #fff;
  }
  .rent-badge-paid { color: ${G.green}; font-weight: 800; font-size: 0.95rem; }
  .rent-badge-unpaid { color: ${G.red}; font-weight: 800; font-size: 0.95rem; }

  /* LEGEND */
  .legend {
    display: flex; gap: 20px; flex-wrap: wrap; margin-bottom: 16px;
    background: #fff; padding: 12px 18px; border-radius: 10px;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06); font-size: 0.8rem;
  }
  .leg-item { display: flex; align-items: center; gap: 7px; font-weight: 500; color: ${G.muted}; }
  .leg-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }

  /* TABLE WRAP */
  .table-wrap {
    border-radius: 14px; overflow: auto;
    box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    max-height: 65vh;
    min-height: 300px;
  }
  table { border-collapse: collapse; width: 100%; min-width: max-content; background: #fff; font-size: 0.83rem; }

    /* HEAD */
    /* Use per-cell sticky positioning so multi-row header rows stack correctly
      and align when the table is horizontally scrolled inside .table-wrap. */
    thead { /* keep as normal flow; stickiness applied to THs */ }
    thead th { position: sticky; z-index: 120; }
    thead tr:first-child th { top: 0; z-index: 125; }
    thead tr:nth-child(2) th { top: 48px; z-index: 124; }

    th {
    padding: 11px 8px; text-align: center; font-weight: 700; font-size: 0.78rem;
    letter-spacing: 0.4px; white-space: nowrap;
    background: linear-gradient(180deg, ${G.navy} 0%, #243658 100%);
    color: #fff; border: 1px solid #2a3f62;
  }
  th.name-col { text-align: left; padding-left: 16px; min-width: 60px; }
  th.friday-head { background: linear-gradient(180deg, ${G.red} 0%, #a00d25 100%) !important; border-color: #8a0b1f !important; }
  th.total-head { background: linear-gradient(180deg, #0d6b31 0%, #0a5226 100%) !important; border-color: #084219 !important; min-width: 90px; }

  /* BODY */
  tbody tr:nth-child(even) td { background: #f7f9fc; }
  tbody tr:hover td { background: #eef3fb !important; }
  tbody tr.summary-row td {
    background: linear-gradient(90deg, ${G.navy}, #243658) !important;
    color: ${G.gold} !important; font-weight: 800; font-size: 0.88rem;
  }
  td {
    padding: 8px 7px; text-align: center; border: 1px solid #e4ecf6;
    vertical-align: middle;
  }
  td.name-cell { text-align: left; padding-left: 14px; font-weight: 700; color: ${G.navy}; min-width: 170px; }
  td.sno-cell { color: ${G.muted}; font-size: 0.78rem; min-width: 36px; }
  td.amount-cell { font-weight: 800; color: ${G.green}; font-size: 0.88rem; }
  td.total-cell {
    font-weight: 800; color: ${G.green}; font-size: 0.9rem;
    background: #f0fdf4 !important;
    border-left: 2px solid #86efac;
  }
  tbody tr:hover td.total-cell { background: #dcfce7 !important; }

  /* CELL INPUTS */
  .cell-sel {
    padding: 4px 5px; border: 1.5px solid #cbd5e1; border-radius: 6px;
    font-size: 0.78rem; font-family: inherit; background: #fff;
    outline: none; cursor: pointer; width: 100%; max-width: 124px;
    transition: border-color 0.15s;
  }
  .cell-sel:focus { border-color: ${G.blue}; }
  .cell-num {
    padding: 4px 5px; border: 1.5px solid #cbd5e1; border-radius: 6px;
    font-size: 0.82rem; font-family: inherit; width: 78px; text-align: center;
    font-weight: 700; outline: none; transition: border-color 0.15s;
  }
  .cell-num:focus { border-color: ${G.green}; }

  /* STATUS BADGE */
  .status-pill {
    display: inline-block; padding: 2px 8px; border-radius: 20px;
    font-size: 0.74rem; font-weight: 700; white-space: nowrap;
  }

  /* PAY MODE TOGGLE */
  .paymode-toggle {
    display: inline-flex; border-radius: 7px; overflow: hidden;
    border: 1.5px solid ${G.border};
  }
  .paymode-btn {
    padding: 4px 9px; font-size: 0.74rem; font-weight: 700; cursor: pointer;
    border: none; font-family: inherit; transition: all 0.15s;
    background: #fff; color: ${G.muted};
  }
  .paymode-btn.active-w { background: ${G.blue}; color: #fff; }
  .paymode-btn.active-m { background: #7c3aed; color: #fff; }

  /* QR SECTION */
  .qr-section {
    margin-top: 56px; text-align: center;
    background: linear-gradient(135deg, ${G.navy} 0%, #1e3a6e 100%);
    border-radius: 20px; padding: 40px 24px 36px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.2);
  }
  .qr-title {
    font-family: 'Syne', sans-serif; font-size: 1.3rem; font-weight: 900;
    color: ${G.gold}; letter-spacing: 2px; margin-bottom: 6px;
  }
  .qr-sub { color: #93b4d8; font-size: 0.88rem; margin-bottom: 24px; }
  .qr-img-wrap {
    display: inline-block; padding: 6px; border-radius: 16px;
    background: linear-gradient(135deg, ${G.gold}, #f5c842);
    box-shadow: 0 8px 24px rgba(240,180,41,0.35);
  }
  .qr-img { display: block; max-width: 100%; height: auto; border-radius: 11px; }
  .qr-name { color: #fff; font-size: 1rem; font-weight: 800; margin-top: 18px; letter-spacing: 0.5px; }
  .phonepe-chip {
    display: inline-flex; align-items: center; gap: 8px;
    background: #5f259f; color: #fff; padding: 8px 22px;
    border-radius: 24px; font-weight: 800; font-size: 0.9rem;
    margin-top: 14px; letter-spacing: 0.5px;
  }
  .phonepe-dot {
    width: 22px; height: 22px; border-radius: 50%; background: rgba(255,255,255,0.2);
    display: flex; align-items: center; justify-content: center;
    font-size: 0.7rem; font-weight: 900;
  }

  @media (max-width: 600px) {
    .topbar-chapter { display: none; }
    .main { padding: 12px 8px 60px; }
    .rent-card { padding: 14px 16px; gap: 12px; }
  }
`

// ─── APP ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [loggedIn, setLoggedIn] = useState(false)
  const [user, setUser] = useState('')
  const [pass, setPass] = useState('')
  const [loginErr, setLoginErr] = useState('')

  const now = new Date()
  const [year, setYear] = useState(now.getFullYear())
  const [month, setMonth] = useState(now.getMonth())

  // data[year][month] = { members: [...], rent: {...} }
  const [allData, setAllData] = useState({})

  const getMonthData = useCallback((y, m) => {
    return allData?.[y]?.[m] ?? null
  }, [allData])

  const monthData = useMemo(() => {
    return getMonthData(year, month) ?? initMonthData(year, month)
  }, [allData, year, month, getMonthData])

  const setMonthData = useCallback((updater) => {
    setAllData(prev => {
      const cur = prev?.[year]?.[month] ?? initMonthData(year, month)
      const next = typeof updater === 'function' ? updater(cur) : updater
      return { ...prev, [year]: { ...(prev[year] ?? {}), [month]: next } }
    })
  }, [year, month])

  const fridays = useMemo(() => getFridays(year, month), [year, month])

  // ── Login ──
  const handleLogin = () => {
    if (user.trim() === 'Chandu' && pass === 'admin@1234') {
      setLoggedIn(true)
      setLoginErr('')
    } else {
      setLoginErr('Invalid username or password.')
    }
  }

  // ── Member field updates ──
  const updatePayMode = (mi, mode) => {
    setMonthData(prev => {
      const members = prev.members.map((m, i) => {
        if (i !== mi) return m
        const newFridays = m.fridays.map(f => ({
          ...f,
          amount: calcAmount(f.status, mode, fridays.length),
        }))
        return { ...m, payMode: mode, fridays: newFridays }
      })
      return { ...prev, members }
    })
  }

  const updateFridayField = (mi, fi, field, value) => {
    setMonthData(prev => {
      const members = prev.members.map((m, i) => {
        if (i !== mi) return m
        const newFridays = m.fridays.map((f, j) => {
          if (j !== fi) return f
          const updated = { ...f, [field]: value }
          if (field === 'status') {
            updated.amount = calcAmount(value, m.payMode, fridays.length)
          }
          return updated
        })
        return { ...m, fridays: newFridays }
      })
      return { ...prev, members }
    })
  }

  const updateRent = (field, value) => {
    setMonthData(prev => ({
      ...prev,
      rent: { ...prev.rent, [field]: value },
    }))
  }

  // ── Export ──
  const exportExcel = () => {
    const wb = XLSX.utils.book_new()
    const { members, rent } = monthData

    const row1 = ['#', 'Member Name', 'Pay Mode']
    const row2 = ['', '', '']
    fridays.forEach((f, i) => {
      row1.push(`Friday ${i + 1} (${fmtDate(f)})`, '', '')
      row2.push('Attendance', 'Amount (₹)', 'Payment Method')
    })
    row1.push('Total Payable (₹)', 'Total Paid (₹)')
    row2.push('', '')

    const rows = [row1, row2]
    let grandPayable = 0, grandPaid = 0

    members.forEach((m, mi) => {
      const row = [mi + 1, MEMBERS[mi], m.payMode === 'weekly' ? 'Weekly' : 'Monthly']
      let totalPayable = 0, totalPaid = 0
      m.fridays.forEach(f => {
        const amt = Number(f.amount) || 0
        totalPayable += amt
        totalPaid += amt
        row.push(f.status.charAt(0).toUpperCase() + f.status.slice(1), amt, f.payMethod.toUpperCase())
      })
      grandPayable += totalPayable
      grandPaid += totalPaid
      row.push(totalPayable, totalPaid)
      rows.push(row)
    })

    const sumRow = ['', 'GRAND TOTAL', '', ...fridays.flatMap(() => ['', '', '']), grandPayable, grandPaid]
    rows.push(sumRow)
    rows.push([])
    rows.push(['VENUE RENT', '', 'Amount (₹)', rent.amount, 'Status', rent.status === 'paid' ? 'PAID ✅' : 'NOT PAID ❌'])

    const ws = XLSX.utils.aoa_to_sheet(rows)
    ws['!cols'] = [{ wch: 4 }, { wch: 24 }, { wch: 12 },
      ...fridays.flatMap(() => [{ wch: 13 }, { wch: 13 }, { wch: 16 }]),
      { wch: 16 }, { wch: 14 }]

    const merges = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },
      { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },
    ]
    fridays.forEach((_, i) => {
      const c = 3 + i * 3
      merges.push({ s: { r: 0, c }, e: { r: 0, c: c + 2 } })
    })
    ws['!merges'] = merges
    XLSX.utils.book_append_sheet(wb, ws, `${MONTHS[month].slice(0,3)} ${year}`)
    XLSX.writeFile(wb, `BNI_Amigos_${MONTHS[month]}_${year}.xlsx`)
  }

  // ── Totals ──
  const memberTotals = useMemo(() => {
    return monthData.members.map(m => {
      const total = m.fridays.reduce((s, f) => s + (Number(f.amount) || 0), 0)
      return total
    })
  }, [monthData])

  const grandTotal = useMemo(() => memberTotals.reduce((a, b) => a + b, 0), [memberTotals])

  // ── Year options ──
  const years = []
  for (let y = now.getFullYear() - 1; y <= now.getFullYear() + 2; y++) years.push(y)

  if (!loggedIn) {
    return (
      <>
        <style>{css}</style>
        <div className="login-wrap">
          <div className="login-card">
            <div className="login-logo">BNI</div>
            <div className="login-sub">Amigos Chapter · Admin Portal</div>
            <div className="login-field">
              <label>Username</label>
              <input
                value={user} onChange={e => setUser(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleLogin()}
                placeholder="Enter username" autoFocus
              />
            </div>
            <div className="login-field">
              <label>Password</label>
              <input
                type="password" value={pass} onChange={e => setPass(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleLogin()}
                placeholder="Enter password"
              />
            </div>
            <button className="login-btn" onClick={handleLogin}>LOGIN →</button>
            {loginErr && <div className="login-err">{loginErr}</div>}
          </div>
        </div>
      </>
    )
  }

  const { members, rent } = monthData

  return (
    <>
      <style>{css}</style>

      {/* TOPBAR */}
      <div className="topbar">
        <div className="topbar-left">
          <span className="topbar-bni">BNI</span>
          <span className="topbar-chapter">Amigos Chapter — Attendance & Payment</span>
        </div>
        <div className="topbar-right">
          <span className="badge-admin">👤 Admin: Chandu</span>
          <button className="btn-logout" onClick={() => { setLoggedIn(false); setUser(''); setPass('') }}>
            Logout
          </button>
        </div>
      </div>

      <div className="main">

        {/* CONTROLS */}
        <div className="controls-row">
          <div className="ctrl-group">
            <span className="ctrl-label">Year</span>
            <select className="ctrl-select" value={year} onChange={e => setYear(Number(e.target.value))}>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
          </div>
          <div className="ctrl-group">
            <span className="ctrl-label">Month</span>
            <select className="ctrl-select" value={month} onChange={e => setMonth(Number(e.target.value))}>
              {MONTHS.map((m, i) => <option key={i} value={i}>{m}</option>)}
            </select>
          </div>
          <span style={{ color: '#64748b', fontSize: '0.82rem', fontWeight: 600 }}>
            📅 {fridays.length} Fridays this month
          </span>
          <button className="btn btn-dl" onClick={exportExcel}>⬇ Download Excel</button>
        </div>

        {/* RENT CARD */}
        <div className="rent-card">
          <div>
            <div className="rent-title">🏛️ Monthly Meeting Venue Rent</div>
            <div className="rent-sub">Amount payable for the venue where this month's meetings are held</div>
          </div>
          <div className="rent-amount-row">
            <span className="rent-rupee">₹</span>
            <input
              type="number" className="rent-input" value={rent.amount}
              onChange={e => updateRent('amount', Number(e.target.value))}
            />
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <span style={{ fontSize: '0.78rem', fontWeight: 700, color: '#64748b', textTransform: 'uppercase' }}>Status:</span>
            <select className="rent-sel" value={rent.status} onChange={e => updateRent('status', e.target.value)}>
              <option value="not_paid">❌ Not Paid</option>
              <option value="paid">✅ Paid</option>
            </select>
          </div>
          <div>
            {rent.status === 'paid'
              ? <span className="rent-badge-paid">✅ PAID — ₹{Number(rent.amount).toLocaleString('en-IN')}</span>
              : <span className="rent-badge-unpaid">❌ PENDING — ₹{Number(rent.amount).toLocaleString('en-IN')}</span>}
          </div>
        </div>

        {/* LEGEND */}
        <div className="legend">
          <strong style={{ color: '#1a2233', marginRight: 4 }}>Key:</strong>
          <div className="leg-item"><div className="leg-dot" style={{ background: '#16a34a' }} /> Present: Weekly ₹800 | Monthly 4F=₹3000, 5F=₹4300</div>
          <div className="leg-item"><div className="leg-dot" style={{ background: '#dc2626' }} /> Absent: ₹1200 (penalty)</div>
          <div className="leg-item"><div className="leg-dot" style={{ background: '#d97706' }} /> Substitute: ₹1000</div>
          <div className="leg-item"><div className="leg-dot" style={{ background: '#7c3aed' }} /> W = Weekly mode | M = Monthly mode</div>
        </div>

        {/* TABLE */}
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th className="name-col" rowSpan={2}>#</th>
                <th className="name-col" rowSpan={2} style={{ minWidth: 160 }}>Member Name</th>
                <th rowSpan={2}>Pay Mode</th>
                {fridays.map((f, i) => (
                  <th key={i} colSpan={3} className="friday-head">
                    Friday {i + 1}<br />
                    <span style={{ fontWeight: 400, fontSize: '0.72rem' }}>{fmtDate(f)}</span>
                  </th>
                ))}
                <th className="total-head" rowSpan={2}>Total<br />Payable (₹)</th>
              </tr>
              <tr>
                {fridays.map((_, i) => (
                  <>
                    <th key={`a${i}`} className="friday-head" style={{ fontSize: '0.72rem' }}>Attendance</th>
                    <th key={`b${i}`} className="friday-head" style={{ fontSize: '0.72rem' }}>Amount (₹)</th>
                    <th key={`c${i}`} className="friday-head" style={{ fontSize: '0.72rem' }}>Payment</th>
                  </>
                ))}
              </tr>
            </thead>
            <tbody>
              {MEMBERS.map((name, mi) => {
                const m = members[mi] ?? initMemberRow(fridays.length)
                return (
                  <tr key={mi}>
                    <td className="sno-cell">{mi + 1}</td>
                    <td className="name-cell">{name}</td>
                    <td>
                      <div className="paymode-toggle">
                        <button
                          className={`paymode-btn ${m.payMode === 'weekly' ? 'active-w' : ''}`}
                          onClick={() => updatePayMode(mi, 'weekly')}
                        >W</button>
                        <button
                          className={`paymode-btn ${m.payMode === 'monthly' ? 'active-m' : ''}`}
                          onClick={() => updatePayMode(mi, 'monthly')}
                        >M</button>
                      </div>
                    </td>
                    {fridays.map((_, fi) => {
                      const fr = m.fridays[fi] ?? { status: 'present', amount: 800, payMethod: 'cash' }
                      const cfg = STATUS_CONFIG[fr.status]
                      return [
                        <td key={`s${fi}`}>
                          <select
                            className="cell-sel"
                            value={fr.status}
                            onChange={e => updateFridayField(mi, fi, 'status', e.target.value)}
                          >
                            {Object.entries(STATUS_CONFIG).map(([k, v]) =>
                              <option key={k} value={k}>{v.label}</option>
                            )}
                          </select>
                        </td>,
                        <td key={`a${fi}`} className="amount-cell">
                          <input
                            type="number"
                            className="cell-num"
                            value={fr.amount}
                            onChange={e => updateFridayField(mi, fi, 'amount', Number(e.target.value))}
                            style={{ color: cfg.color }}
                          />
                        </td>,
                        <td key={`p${fi}`}>
                          <select
                            className="cell-sel"
                            value={fr.payMethod}
                            onChange={e => updateFridayField(mi, fi, 'payMethod', e.target.value)}
                          >
                            <option value="cash">💵 Cash</option>
                            <option value="upi">📲 UPI</option>
                          </select>
                        </td>,
                      ]
                    })}
                    <td className="total-cell">
                      ₹{memberTotals[mi]?.toLocaleString('en-IN') ?? 0}
                    </td>
                  </tr>
                )
              })}
              <tr className="summary-row">
                <td></td>
                <td style={{ paddingLeft: 14 }}>GRAND TOTAL</td>
                <td></td>
                {fridays.flatMap((_, i) => [<td key={`s${i}`}></td>, <td key={`a${i}`}></td>, <td key={`p${i}`}></td>])}
                <td>₹{grandTotal.toLocaleString('en-IN')}</td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* QR SECTION */}
        <div className="qr-section">
          <div className="qr-title">💳 PAYMENT — PHONEPE / UPI</div>
          <div className="qr-sub">Scan this QR code using PhonePe or any UPI app to make your payment</div>
          <div className="qr-img-wrap">
            <img src={QR_BASE64} alt="PhonePe QR Code" className="qr-img" />
          </div>
          <div className="qr-name">Mr. Vobbilisetty Chandra Krishna</div>
          <div>
            <div className="phonepe-chip">
              <div className="phonepe-dot">Pe</div>
              PhonePe Accepted Here
            </div>
          </div>
        </div>

      </div>
    </>
  )
}
