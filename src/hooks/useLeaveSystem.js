import { useState, useEffect, useMemo } from 'react';

const parseISO = (str) => new Date(str);
const getYear = (date) => date.getFullYear();
const differenceInDays = (d1, d2) => {
  const diffTime = Math.abs(d1 - d2);
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
};

const STORAGE_KEY = 'maki_leave_system_data';
// 您可以在下方引號內填入您的 Google Apps Script 網址，這樣所有人打開網頁就會自動連線
const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbyWzMNXzBTwOgXqHibTFRvmjKfzZ38TZG2iVcEl7dI_ltitOnW33TTr_vOt5QgG1WmoRA/exec'; 

export const useLeaveSystem = () => {
  const [apiUrl, setApiUrl] = useState(DEFAULT_API_URL);
  const [isSyncing, setIsSyncing] = useState(false);
  const [data, setData] = useState(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    return saved ? JSON.parse(saved) : { salespersons: [], records: [] };
  });



  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    if (apiUrl) {
      const timeoutId = setTimeout(() => saveToGAS(), 1000); // Debounce save
      return () => clearTimeout(timeoutId);
    }
  }, [data]);



  // Auto-sync on mount and window focus
  useEffect(() => {
    if (apiUrl) {
      syncData();
      const handleFocus = () => syncData();
      window.addEventListener('focus', handleFocus);
      return () => window.removeEventListener('focus', handleFocus);
    }
  }, [apiUrl]);

  const syncData = async (targetUrl = apiUrl) => {
    if (!targetUrl || isSyncing) return;
    setIsSyncing(true);
    try {
      const res = await fetch(targetUrl + "?action=getData");
      const remoteData = await res.json();
      if (remoteData.salespersons) {
        // Only update if data is actually different to avoid infinite loops
        setData(prev => {
          if (JSON.stringify(prev) === JSON.stringify(remoteData)) return prev;
          return remoteData;
        });
      }
    } catch (e) {
      console.error("Auto-sync failed:", e);
    } finally {
      setIsSyncing(false);
    }
  };

  const saveToGAS = async () => {
    if (!apiUrl) return;
    try {
      await fetch(apiUrl, {
        method: 'POST',
        body: JSON.stringify({ action: 'saveData', data }),
      });
    } catch (e) {
      console.warn("Auto-save failed:", e);
    }
  };

  // Actions
  const addSalesperson = (person) => {
    setData(prev => ({
      ...prev,
      salespersons: [...prev.salespersons, { ...person, id: crypto.randomUUID() }]
    }));
  };

  const updateSalesperson = (id, updates) => {
    setData(prev => ({
      ...prev,
      salespersons: prev.salespersons.map(p => p.id === id ? { ...p, ...updates } : p)
    }));
  };

  const deleteSalesperson = (id) => {
    setData(prev => ({
      ...prev,
      salespersons: prev.salespersons.filter(p => p.id !== id),
      records: prev.records.filter(r => r.salespersonId !== id)
    }));
  };

  const addRecord = (record) => {
    setData(prev => ({
      ...prev,
      records: [...prev.records, { ...record, id: crypto.randomUUID() }]
    }));
  };

  const importData = (rawRecords) => {
    setData(prev => {
      const newSalespersons = [...prev.salespersons];
      const newRecords = [...prev.records];

      rawRecords.forEach(item => {
        // Find or create salesperson
        let person = newSalespersons.find(p => p.name.trim() === item.name.trim());
        if (!person) {
          person = { id: crypto.randomUUID(), name: item.name.trim(), department: item.department || '', note: '' };
          newSalespersons.push(person);
        }

        // Add record
        newRecords.push({
          id: crypto.randomUUID(),
          salespersonId: person.id,
          startDate: item.startDate,
          endDate: item.endDate,
          location: item.location || '',
          note: item.note || ''
        });
      });

      return { salespersons: newSalespersons, records: newRecords };
    });
  };

  const updateRecord = (id, updates) => {
    setData(prev => ({
      ...prev,
      records: prev.records.map(r => r.id === id ? { ...r, ...updates } : r)
    }));
  };

  const deleteRecord = (id) => {
    setData(prev => ({
      ...prev,
      records: prev.records.filter(r => r.id !== id)
    }));
  };

  // Calculation Logic
  const processedData = useMemo(() => {
    const { salespersons, records } = data;
    
    // Group records by salesperson and year
    const results = {};

    salespersons.forEach(person => {
      results[person.id] = {
        info: person,
        years: {}
      };

      const personRecords = records.filter(r => r.salespersonId === person.id);
      
      // Get all unique years
      const years = [...new Set(personRecords.map(r => getYear(parseISO(r.startDate))))].sort((a, b) => b - a);
      
      years.forEach(year => {
        const yearRecords = personRecords
          .filter(r => getYear(parseISO(r.startDate)) === year)
          .sort((a, b) => a.startDate.localeCompare(b.startDate));

        let runningRemainder = 0;
        const calculatedRecords = yearRecords.map(record => {
          const start = parseISO(record.startDate);
          const end = parseISO(record.endDate);
          const tripDays = differenceInDays(end, start) + 1;
          const weeks = tripDays / 7;
          const generatedLeave = weeks * 0.8;
          
          const totalWithCarryover = generatedLeave + runningRemainder;
          const awarded = Math.floor(totalWithCarryover + 0.000001); // Small epsilon for floating point issues
          const remainder = totalWithCarryover - awarded;
          
          const result = {
            ...record,
            tripDays,
            weeks: parseFloat(weeks.toFixed(4)),
            generatedLeave: parseFloat(generatedLeave.toFixed(2)),
            prevRemainder: parseFloat(runningRemainder.toFixed(2)),
            awarded,
            remainder: parseFloat(remainder.toFixed(2))
          };

          runningRemainder = remainder;
          return result;
        });

        results[person.id].years[year] = {
          records: calculatedRecords,
          totalTripDays: calculatedRecords.reduce((sum, r) => sum + r.tripDays, 0),
          totalAwarded: calculatedRecords.reduce((sum, r) => sum + r.awarded, 0),
          finalRemainder: parseFloat(runningRemainder.toFixed(2)),
          tripCount: calculatedRecords.length
        };
      });
    });

    return results;
  }, [data]);

  const allYears = useMemo(() => {
    const years = new Set();
    data.records.forEach(r => years.add(getYear(parseISO(r.startDate))));
    return years.size > 0 ? [...years].sort((a, b) => b - a) : [new Date().getFullYear()];
  }, [data.records]);

  return {
    salespersons: data.salespersons,
    records: data.records,
    processedData,
    allYears,
    isSyncing,
    apiUrl,
    actions: {
      addSalesperson,
      updateSalesperson,
      deleteSalesperson,
      addRecord,
      updateRecord,
      deleteRecord,
      importData,
      setApiUrl,
      syncData
    }
  };
};
