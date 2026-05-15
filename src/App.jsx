import React, { useState, useEffect } from 'react';
import { useLeaveSystem } from './hooks/useLeaveSystem';
import { exportToExcel, exportPersonDetails } from './utils/excelExport';
import * as XLSX from 'xlsx';
import { 
  Users, 
  Calendar, 
  Plus, 
  Download, 
  Upload,
  Trash2, 
  Edit2, 
  ChevronRight, 
  X,
  TrendingUp,
  Award,
  History,
  Plane,
  Settings,
  RefreshCw,
  Cloud,
  CloudOff,
  Languages
} from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import { clsx } from 'clsx';

const formatDate = (date) => {
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  let month = '' + (d.getMonth() + 1);
  let day = '' + d.getDate();
  const year = d.getFullYear();
  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;
  return [year, month, day].join('-');
};

const translations = {
  zh: {
    title: 'Special Paid Leave',
    subtitle: '不扣薪事假計算系統',
    year: '年度',
    exportSummary: '匯出總覽',
    importData: '匯出資料', // Actually "匯入資料" in previous code, fixing it to "匯入資料"
    addPerson: '新增業務',
    statsTotalSales: '業務總數',
    statsTotalTrips: '年度出差總次數',
    statsTotalDays: '年度出差總天數',
    statsTotalAwarded: '已核給特休總天數',
    statsSalesUnit: '位業務人員',
    statsTripsUnit: '次出差紀錄',
    statsDaysUnit: '個工作日',
    statsAwardedUnit: '天不扣薪假',
    tableTitle: '年度業務統計',
    colName: '業務姓名',
    colDept: '部門 / 區域',
    colTrips: '出差次數',
    colDays: '出差總天數',
    colAwarded: '已核給 Leave',
    colRemainder: '小數餘額',
    colActions: '操作',
    noData: '尚無業務人員資料，請點擊右上方「新增業務」開始。',
    syncing: '同步中...',
    connected: '已連線雲端',
    localMode: '本地模式',
    personDetails: '年度出差明細',
    exportDetails: '匯出個人明細',
    addRecord: '新增出差紀錄',
    colPeriod: '日期區間',
    colGenLeave: '產生天數',
    colPrevRem: '前次餘額',
    colNote: '地點/備註',
    close: '關閉',
    editPerson: '編輯業務資料',
    editRecord: '編輯出差紀錄',
    save: '儲存',
    cancel: '取消',
    saveRecord: '儲存紀錄',
    settingsTitle: '雲端同步設定',
    settingsDesc: '輸入 Google Apps Script 網頁應用程式網址，即可實現手機與電腦資料同步。',
    syncNow: '立即同步',
    done: '完成',
    confirmDeletePerson: '確定要刪除該業務及其所有紀錄嗎？',
    confirmDeleteRecord: '確定刪除此紀錄？',
    importSuccess: '成功匯入 {n} 筆紀錄！',
    importFailed: '匯入失敗，請確保檔案格式正確。',
    importInvalid: '找不到有效的紀錄，請檢查 Excel 欄位名稱是否正確。',
    startDate: '開始日期',
    endDate: '結束日期',
    location: '出差地點',
    note: '備註',
    deptPlaceholder: '例如：北美業務部',
    locPlaceholder: '例如：美國舊金山',
    notePlaceholder: '選填備註資訊',
    recordNotePlaceholder: '出差事由或其他備註'
  },
  en: {
    title: 'Special Paid Leave',
    subtitle: 'Business Trip Leave Calculator',
    year: 'Year',
    exportSummary: 'Export Summary',
    importData: 'Import Data',
    addPerson: 'Add Personnel',
    statsTotalSales: 'Total Sales',
    statsTotalTrips: 'Total Trips (YTD)',
    statsTotalDays: 'Total Days (YTD)',
    statsTotalAwarded: 'Total Leave Awarded',
    statsSalesUnit: 'persons',
    statsTripsUnit: 'trips',
    statsDaysUnit: 'days',
    statsAwardedUnit: 'days',
    tableTitle: 'Annual Statistics',
    colName: 'Name',
    colDept: 'Dept / Region',
    colTrips: 'Trips',
    colDays: 'Total Days',
    colAwarded: 'Awarded Leave',
    colRemainder: 'Remainder',
    colActions: 'Actions',
    noData: 'No data yet. Click "Add Personnel" to start.',
    syncing: 'Syncing...',
    connected: 'Cloud Connected',
    localMode: 'Local Mode',
    personDetails: 'Trip Details',
    exportDetails: 'Export Details',
    addRecord: 'Add Record',
    colPeriod: 'Date Period',
    colGenLeave: 'Generated',
    colPrevRem: 'Prev Remainder',
    colNote: 'Location / Note',
    close: 'Close',
    editPerson: 'Edit Personnel',
    editRecord: 'Edit Record',
    save: 'Save',
    cancel: 'Cancel',
    saveRecord: 'Save Record',
    settingsTitle: 'Cloud Sync Settings',
    settingsDesc: 'Enter your Google Apps Script URL to sync between mobile and desktop.',
    syncNow: 'Sync Now',
    done: 'Done',
    confirmDeletePerson: 'Delete this personnel and all records?',
    confirmDeleteRecord: 'Delete this record?',
    importSuccess: 'Successfully imported {n} records!',
    importFailed: 'Import failed. Check file format.',
    importInvalid: 'No valid records found. Check Excel headers.',
    startDate: 'Start Date',
    endDate: 'End Date',
    location: 'Location',
    note: 'Note',
    deptPlaceholder: 'e.g., Sales Dept',
    locPlaceholder: 'e.g., San Francisco, USA',
    notePlaceholder: 'Optional notes',
    recordNotePlaceholder: 'Purpose or details'
  }
};

const App = () => {
  const { salespersons, records, processedData, allYears, isSyncing, apiUrl, actions } = useLeaveSystem();
  const [lang, setLang] = useState(() => localStorage.getItem('maki_lang') || 'zh');
  const [selectedYear, setSelectedYear] = useState(new Date().getFullYear());
  const [view, setView] = useState('dashboard');
  const [selectedPersonId, setSelectedPersonId] = useState(null);

  const t = (key) => translations[lang][key] || key;

  useEffect(() => {
    localStorage.setItem('maki_lang', lang);
  }, [lang]);

  // Modals state
  const [showPersonModal, setShowPersonModal] = useState(false);
  const [showRecordModal, setShowRecordModal] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [editingPerson, setEditingPerson] = useState(null);
  const [editingRecord, setEditingRecord] = useState(null);

  // Stats for the selected year
  const yearStats = Object.values(processedData).reduce((acc, p) => {
    const yearData = p.years[selectedYear];
    if (yearData) {
      acc.totalTrips += yearData.tripCount;
      acc.totalDays += yearData.totalTripDays;
      acc.totalAwarded += yearData.totalAwarded;
    }
    return acc;
  }, { totalTrips: 0, totalDays: 0, totalAwarded: 0 });

  const handleExportAll = () => {
    const data = salespersons.map(p => {
      const stats = processedData[p.id]?.years[selectedYear] || { tripCount: 0, totalTripDays: 0, totalAwarded: 0, finalRemainder: 0 };
      return {
        '業務姓名': p.name,
        '年度': selectedYear,
        '出差次數': stats.tripCount,
        '出差總天數': stats.totalTripDays,
        '已核給特休': stats.totalAwarded,
        '小數餘額': stats.finalRemainder,
        '部門': p.department || '',
        '備註': p.note || ''
      };
    });
    exportToExcel(data, `Sales_Leave_Summary_${selectedYear}.xlsx`);
  };

  const handleExportPerson = (personId) => {
    const person = salespersons.find(p => p.id === personId);
    const yearData = processedData[personId]?.years[selectedYear];
    if (person && yearData) {
      exportPersonDetails(person.name, selectedYear, yearData.records);
    }
  };

  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);
        
        // Expected columns: 業務姓名, 開始日期, 結束日期, 地點, 備註, 部門
        const formatted = data.map(row => ({
          name: row['業務姓名']?.toString() || '',
          startDate: row['開始日期']?.toString() || '',
          endDate: row['結束日期']?.toString() || '',
          location: row['地點']?.toString() || row['出差地點']?.toString() || '',
          note: row['備註']?.toString() || '',
          department: row['部門']?.toString() || ''
        })).filter(r => r.name && r.startDate && r.endDate);

        if (formatted.length === 0) {
          alert('找不到有效的紀錄，請檢查 Excel 欄位名稱是否正確（業務姓名, 開始日期, 結束日期）。');
          return;
        }

        actions.importData(formatted);
        alert(`成功匯入 ${formatted.length} 筆紀錄！`);
      } catch (err) {
        console.error(err);
        alert('匯入失敗，請確保檔案格式正確。');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = ''; // Reset input
  };

  return (
    <div className="app-container">
      {/* Header */}
      <header className="animate-fade-in" style={{ marginBottom: '2.5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <h1 style={{ fontSize: '2.5rem', fontWeight: 800, letterSpacing: '-0.02em', background: 'linear-gradient(to right, #10b981, #3b82f6)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
            {t('title')}
          </h1>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', marginTop: '0.25rem' }}>
            <p style={{ color: 'var(--text-secondary)' }}>{t('subtitle')}</p>
            {apiUrl ? (
              <span style={{ fontSize: '0.75rem', color: isSyncing ? 'var(--accent-color)' : 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                <Cloud size={14} /> {isSyncing ? t('syncing') : t('connected')}
              </span>
            ) : (
              <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                <CloudOff size={14} /> {t('localMode')}
              </span>
            )}
          </div>
        </div>
        
        <div className="header-actions">
          <button className="secondary" style={{ padding: '0.75rem' }} onClick={() => setLang(lang === 'zh' ? 'en' : 'zh')}>
            <Languages size={20} />
          </button>
          <button className="secondary" style={{ padding: '0.75rem' }} onClick={() => setShowSettingsModal(true)}>
            <Settings size={20} />
          </button>
          <select 
            value={selectedYear} 
            onChange={(e) => setSelectedYear(parseInt(e.target.value))}
            className="year-select"
          >
            {allYears.map(y => <option key={y} value={y}>{y} {t('year')}</option>)}
          </select>
          <button className="secondary" onClick={handleExportAll}>
            <Download size={18} /> <span className="hide-mobile">{t('exportSummary')}</span>
          </button>
          <label className="secondary import-label">
            <Upload size={18} /> <span className="hide-mobile">{t('importData')}</span>
            <input type="file" accept=".xlsx, .xls" style={{ display: 'none' }} onChange={handleImport} />
          </label>
          <button className="primary" onClick={() => setShowPersonModal(true)}>
            <Plus size={18} /> <span className="hide-mobile">{t('addPerson')}</span>
          </button>
        </div>
      </header>

      {/* Stats Overview */}
      <section className="stats-grid">
        <StatsCard icon={<Users size={24} />} title={t('statsTotalSales')} value={salespersons.length} label={t('statsSalesUnit')} />
        <StatsCard icon={<Plane size={24} />} title={t('statsTotalTrips')} value={yearStats.totalTrips} label={t('statsTripsUnit')} />
        <StatsCard icon={<TrendingUp size={24} />} title={t('statsTotalDays')} value={yearStats.totalDays} label={t('statsDaysUnit')} />
        <StatsCard icon={<Award size={24} />} title={t('statsTotalAwarded')} value={yearStats.totalAwarded} label={t('statsAwardedUnit')} />
      </section>

      {/* Main Content */}
      <main className="glass-card animate-fade-in" style={{ padding: '1.5rem', minHeight: '400px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.5rem' }}>
          <h2 style={{ fontSize: '1.25rem', display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
            <History size={20} color="var(--accent-color)" /> {selectedYear} {t('tableTitle')}
          </h2>
        </div>

        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>{t('colName')}</th>
                <th>{t('colDept')}</th>
                <th>{t('colTrips')}</th>
                <th>{t('colDays')}</th>
                <th>{t('colAwarded')}</th>
                <th>{t('colRemainder')}</th>
                <th style={{ textAlign: 'right' }}>{t('colActions')}</th>
              </tr>
            </thead>
            <tbody>
              {salespersons.length === 0 ? (
                <tr>
                  <td colSpan="7" style={{ textAlign: 'center', padding: '4rem', color: 'var(--text-secondary)' }}>
                    {t('noData')}
                  </td>
                </tr>
              ) : (
                salespersons.map(person => {
                  const stats = processedData[person.id]?.years[selectedYear] || { tripCount: 0, totalTripDays: 0, totalAwarded: 0, finalRemainder: 0 };
                  return (
                    <tr key={person.id} style={{ cursor: 'pointer' }} onClick={() => { setSelectedPersonId(person.id); setView('person-details'); }}>
                      <td style={{ fontWeight: 600 }}>{person.name}</td>
                      <td style={{ color: 'var(--text-secondary)' }}>{person.department || '—'}</td>
                      <td>{stats.tripCount}</td>
                      <td>{stats.totalTripDays}</td>
                      <td>
                        <span className="badge badge-accent">{stats.totalAwarded} {t('statsAwardedUnit')}</span>
                      </td>
                      <td style={{ fontVariantNumeric: 'tabular-nums' }}>{stats.finalRemainder.toFixed(2)}</td>
                      <td style={{ textAlign: 'right' }}>
                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.5rem' }} onClick={e => e.stopPropagation()}>
                          <button className="secondary" style={{ padding: '0.5rem' }} onClick={() => { setEditingPerson(person); setShowPersonModal(true); }}>
                            <Edit2 size={16} />
                          </button>
                          <button className="secondary" style={{ padding: '0.5rem', color: 'var(--danger-color)' }} onClick={() => { if(confirm(t('confirmDeletePerson'))) actions.deleteSalesperson(person.id); }}>
                            <Trash2 size={16} />
                          </button>
                          <ChevronRight size={20} color="var(--text-secondary)" style={{ marginLeft: '0.5rem' }} />
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </main>

      {/* Person Details Drawer/Overlay */}
      <AnimatePresence>
        {view === 'person-details' && selectedPersonId && (
          <motion.div 
            initial={{ x: '100%' }}
            animate={{ x: 0 }}
            exit={{ x: '100%' }}
            transition={{ type: 'spring', damping: 25, stiffness: 200 }}
            style={{ position: 'fixed', top: 0, right: 0, height: '100vh', width: 'min(90%, 1000px)', background: '#0a0a0a', borderLeft: '1px solid var(--surface-border)', zIndex: 100, padding: '2rem', overflowY: 'auto', boxShadow: '-20px 0 50px rgba(0,0,0,0.5)' }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '2rem' }}>
              <div>
                <button className="secondary" onClick={() => setView('dashboard')} style={{ marginBottom: '1rem', padding: '0.5rem 1rem' }}>
                  <X size={18} /> {t('close')}
                </button>
                <h2 style={{ fontSize: '2rem' }}>{salespersons.find(p => p.id === selectedPersonId)?.name}</h2>
                <p style={{ color: 'var(--text-secondary)' }}>{selectedYear} {t('personDetails')}</p>
              </div>
              <div style={{ display: 'flex', gap: '1rem' }}>
                <button className="secondary" onClick={() => handleExportPerson(selectedPersonId)}>
                  <Download size={18} /> {t('exportDetails')}
                </button>
                <button className="primary" onClick={() => { setEditingRecord(null); setShowRecordModal(true); }}>
                  <Plus size={18} /> {t('addRecord')}
                </button>
              </div>
            </div>

            <div className="glass-card" style={{ padding: '1rem' }}>
              <div className="table-wrapper">
                <table className="details-table">
                <thead>
                  <tr>
                    <th>{t('colPeriod')}</th>
                    <th>{t('statsDaysUnit')}</th>
                    <th>{t('statsTripsUnit')}</th> {/* This was weeks, changing to a generic unit or translated */}
                    <th>{t('colGenLeave')}</th>
                    <th>{t('colPrevRem')}</th>
                    <th>{t('colAwarded')}</th>
                    <th>{t('colRemainder')}</th>
                    <th>{t('colNote')}</th>
                    <th>{t('colActions')}</th>
                  </tr>
                </thead>
                <tbody>
                  {(processedData[selectedPersonId]?.years[selectedYear]?.records || []).map(record => (
                    <tr key={record.id}>
                      <td style={{ fontSize: '0.875rem' }}>
                        {record.startDate} <br/> <span style={{ color: 'var(--text-secondary)' }}>{t('endDate')}: {record.endDate}</span>
                      </td>
                      <td style={{ fontWeight: 600 }}>{record.tripDays}</td>
                      <td style={{ color: 'var(--text-secondary)' }}>{record.weeks}</td>
                      <td>{record.generatedLeave.toFixed(2)}</td>
                      <td>{record.prevRemainder.toFixed(2)}</td>
                      <td>
                        <span className="badge badge-accent">{record.awarded} {t('statsAwardedUnit')}</span>
                      </td>
                      <td style={{ fontWeight: 600 }}>{record.remainder.toFixed(2)}</td>
                      <td style={{ maxWidth: '200px', fontSize: '0.875rem' }}>
                        <div style={{ fontWeight: 500 }}>{record.location}</div>
                        <div style={{ color: 'var(--text-secondary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{record.note}</div>
                      </td>
                      <td>
                        <div style={{ display: 'flex', gap: '0.25rem' }}>
                          <button className="secondary" style={{ padding: '0.4rem' }} onClick={() => { setEditingRecord(record); setShowRecordModal(true); }}>
                            <Edit2 size={14} />
                          </button>
                          <button className="secondary" style={{ padding: '0.4rem', color: 'var(--danger-color)' }} onClick={() => { if(confirm(t('confirmDeleteRecord'))) actions.deleteRecord(record.id); }}>
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </motion.div>
        )}
      </AnimatePresence>

      {/* Person Modal */}
      {showPersonModal && (
        <Modal 
          title={editingPerson ? t('editPerson') : t('addPerson')} 
          onClose={() => { setShowPersonModal(false); setEditingPerson(null); }}
        >
          <PersonForm 
            initialData={editingPerson} 
            t={t}
            onSubmit={(data) => {
              if (editingPerson) actions.updateSalesperson(editingPerson.id, data);
              else actions.addSalesperson(data);
              setShowPersonModal(false);
              setEditingPerson(null);
            }} 
            onCancel={() => { setShowPersonModal(false); setEditingPerson(null); }}
          />
        </Modal>
      )}

      {/* Record Modal */}
      {showRecordModal && (
        <Modal 
          title={editingRecord ? t('editRecord') : t('addRecord')} 
          onClose={() => { setShowRecordModal(false); setEditingRecord(null); }}
        >
          <RecordForm 
            initialData={editingRecord}
            personId={selectedPersonId}
            t={t}
            onSubmit={(data) => {
              if (editingRecord) actions.updateRecord(editingRecord.id, data);
              else actions.addRecord({ ...data, salespersonId: selectedPersonId });
              setShowRecordModal(false);
              setEditingRecord(null);
            }} 
            onCancel={() => { setShowRecordModal(false); setEditingRecord(null); }}
          />
        </Modal>
      )}

      {/* Settings Modal */}
      {showSettingsModal && (
        <Modal 
          title={t('settingsTitle')} 
          onClose={() => setShowSettingsModal(false)}
        >
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
            <p style={{ fontSize: '0.875rem', color: 'var(--text-secondary)' }}>
              {t('settingsDesc')}
            </p>
            <div>

            <div style={{ display: 'flex', gap: '1rem' }}>
              <button className="secondary" style={{ flex: 1 }} onClick={() => actions.syncData()}>
                <RefreshCw size={18} className={isSyncing ? 'animate-spin' : ''} /> {t('syncNow')}
              </button>
              <button className="primary" style={{ flex: 1 }} onClick={() => setShowSettingsModal(false)}>
                {t('done')}
              </button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
};

const StatsCard = ({ icon, title, value, label }) => (
  <div className="glass-card" style={{ padding: '1.5rem', display: 'flex', alignItems: 'center', gap: '1.25rem' }}>
    <div style={{ background: 'var(--accent-glow)', padding: '1rem', borderRadius: '1rem', color: 'var(--accent-color)' }}>
      {icon}
    </div>
    <div>
      <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem', marginBottom: '0.25rem' }}>{title}</p>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: '0.5rem' }}>
        <h3 style={{ fontSize: '1.75rem', fontWeight: 700 }}>{value}</h3>
        <span style={{ fontSize: '0.875rem', color: 'var(--text-secondary)' }}>{label}</span>
      </div>
    </div>
  </div>
);

const Modal = ({ title, children, onClose }) => (
  <div style={{ position: 'fixed', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 200, padding: '1rem' }}>
    <motion.div 
      initial={{ opacity: 0 }} 
      animate={{ opacity: 1 }} 
      exit={{ opacity: 0 }}
      onClick={onClose}
      style={{ position: 'absolute', inset: 0, background: 'rgba(0,0,0,0.8)', backdropFilter: 'blur(4px)' }} 
    />
    <motion.div 
      initial={{ scale: 0.9, opacity: 0, y: 20 }} 
      animate={{ scale: 1, opacity: 1, y: 0 }}
      exit={{ scale: 0.9, opacity: 0, y: 20 }}
      className="glass-card" 
      style={{ width: '100%', maxWidth: '500px', padding: '2rem', position: 'relative', zIndex: 201 }}
    >
      <h3 style={{ fontSize: '1.5rem', marginBottom: '1.5rem' }}>{title}</h3>
      {children}
    </motion.div>
  </div>
);

const PersonForm = ({ initialData, t, onSubmit, onCancel }) => {
  const [formData, setFormData] = useState(initialData || { name: '', department: '', note: '' });
  return (
    <form onSubmit={(e) => { e.preventDefault(); onSubmit(formData); }} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
      <div>
        <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('colName')}</label>
        <input required value={formData.name} onChange={e => setFormData({ ...formData, name: e.target.value })} placeholder="Asher" />
      </div>
      <div>
        <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('colDept')}</label>
        <input value={formData.department} onChange={e => setFormData({ ...formData, department: e.target.value })} placeholder={t('deptPlaceholder')} />
      </div>
      <div>
        <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('note')}</label>
        <textarea value={formData.note} onChange={e => setFormData({ ...formData, note: e.target.value })} placeholder={t('notePlaceholder')} />
      </div>
      <div style={{ display: 'flex', gap: '1rem', marginTop: '0.5rem' }}>
        <button type="button" className="secondary" onClick={onCancel} style={{ flex: 1 }}>{t('cancel')}</button>
        <button type="submit" className="primary" style={{ flex: 2 }}>{t('save')}</button>
      </div>
    </form>
  );
};

const RecordForm = ({ initialData, t, onSubmit, onCancel }) => {
  const [formData, setFormData] = useState(initialData || { startDate: formatDate(new Date()), endDate: formatDate(new Date()), location: '', note: '' });
  return (
    <form onSubmit={(e) => { e.preventDefault(); onSubmit(formData); }} style={{ display: 'flex', flexDirection: 'column', gap: '1.25rem' }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
        <div>
          <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('startDate')}</label>
          <input type="date" required value={formData.startDate} onChange={e => setFormData({ ...formData, startDate: e.target.value })} />
        </div>
        <div>
          <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('endDate')}</label>
          <input type="date" required value={formData.endDate} onChange={e => setFormData({ ...formData, endDate: e.target.value })} />
        </div>
      </div>
      <div>
        <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('location')}</label>
        <input value={formData.location} onChange={e => setFormData({ ...formData, location: e.target.value })} placeholder={t('locPlaceholder')} />
      </div>
      <div>
        <label style={{ display: 'block', fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '0.5rem' }}>{t('note')}</label>
        <textarea value={formData.note} onChange={e => setFormData({ ...formData, note: e.target.value })} placeholder={t('recordNotePlaceholder')} />
      </div>
      <div style={{ display: 'flex', gap: '1rem', marginTop: '0.5rem' }}>
        <button type="button" className="secondary" onClick={onCancel} style={{ flex: 1 }}>{t('cancel')}</button>
        <button type="submit" className="primary" style={{ flex: 2 }}>{t('saveRecord')}</button>
      </div>
    </form>
  );
};

export default App;
