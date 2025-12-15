import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';

const TEST_IDS = [
  'TMPDScb32d04b64d94a9', 'TMPDS4abdb524d673492', 'TMPDS5254acb93dbe46c',
  'TMPDSa2686c826a28485', 'TMPDS6a4757e6a3c34cc', 'TMPDSc53c81cb026f488',
  'TMPDS067d9b743d17463', 'TMPDS43098c59653c486', 'TMPDS21c02640426e436',
  'TMPDS8b09cd30f54e476', 'TMPDSd27bf78fb8e546a', 'TMPDSd5034a6fbad64be',
  'TMPDS77970861beae492', 'TMPDS28c045ff094843a', 'TMPDS4ccba6a2a15040e',
  'TMPDS731a0fb561354e0', 'TMPDS9fb6acec8fe14b8', 'TMPDSa9f21742c6e1b84',
  'TMPDSe5a4afa77d6346f', 'TMPDS1e7083124613423', 'TMPDSabb9d72cecd244d',
];

export default function App() {
  const [data, setData] = useState([]);
  const [fileName, setFileName] = useState('');
  const [mainTab, setMainTab] = useState('earn');
  const [useSubTab, setUseSubTab] = useState('company');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('');
  const [showCanceled, setShowCanceled] = useState(false);
  const [expandedUser, setExpandedUser] = useState(null);
  const [expandedEarn, setExpandedEarn] = useState(null);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);
        setData(jsonData);
        
        if (jsonData.length > 0) {
          const firstDate = jsonData[0]['처리일'] || '';
          const match = firstDate.match(/(\d{4})\/(\d{2})/);
          if (match) {
            setSelectedMonth(`${match[1]}-${match[2]}`);
          }
        }
      } catch (err) {
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const realData = useMemo(() => {
    return data.filter(item => !TEST_IDS.includes(item['고객ID']));
  }, [data]);

  const validData = useMemo(() => {
    if (showCanceled) return realData;
    return realData.filter(item => item['주문상태'] !== '취소완료');
  }, [realData, showCanceled]);

  const canceledCount = useMemo(() => {
    return realData.filter(item => item['주문상태'] === '취소완료').length;
  }, [realData]);

  const monthlyData = useMemo(() => {
    if (!selectedMonth) return validData;
    return validData.filter(row => {
      const date = row['처리일'] || '';
      return date.startsWith(selectedMonth.replace('-', '/'));
    });
  }, [validData, selectedMonth]);

  const availableMonths = useMemo(() => {
    const months = new Set();
    validData.forEach(row => {
      const date = row['처리일'] || '';
      const match = date.match(/(\d{4})\/(\d{2})/);
      if (match) {
        months.add(`${match[1]}-${match[2]}`);
      }
    });
    return Array.from(months).sort().reverse();
  }, [validData]);

  const earnData = useMemo(() => monthlyData.filter(r => r['타입'] !== '사용'), [monthlyData]);
  const useData = useMemo(() => monthlyData.filter(r => r['타입'] === '사용'), [monthlyData]);

  const carryoverPoint = useMemo(() => {
    if (!selectedMonth) return 0;
    const prevData = validData.filter(row => {
      const date = row['처리일'] || '';
      const rowMonth = date.substring(0, 7).replace('/', '-');
      return rowMonth < selectedMonth;
    });
    if (prevData.length === 0) return 0;
    
    // 사용자별 가장 마지막 거래의 토탈포인트
    const userLastTotal = new Map();
    prevData
      .sort((a, b) => (a['처리일'] || '').localeCompare(b['처리일'] || ''))
      .forEach(row => {
        userLastTotal.set(row['고객ID'], Number(row['토탈포인트']) || 0);
      });
    return Array.from(userLastTotal.values()).reduce((sum, v) => sum + v, 0);
  }, [validData, selectedMonth]);

  const monthlyTotals = useMemo(() => {
    let totalUsed = 0;
    let totalEarned = 0;
    monthlyData.forEach(row => {
      const point = Number(row['포인트']) || 0;
      if (row['타입'] === '사용') {
        totalUsed += point;
      } else {
        totalEarned += point;
      }
    });
    return {
      used: totalUsed,
      earned: totalEarned,
      carryover: carryoverPoint,
      balance: carryoverPoint + totalEarned + totalUsed
    };
  }, [monthlyData, carryoverPoint]);

  const earnByType = useMemo(() => {
    const map = new Map();
    earnData.forEach(row => {
      let memo = (row['관리자메모'] || '(메모없음)').trim().replace(/\n/g, '');
      const point = Number(row['포인트']) || 0;
      const userId = row['고객ID'];
      const date = (row['처리일'] || '').substring(0, 10);
      
      if (!map.has(memo)) {
        map.set(memo, { memo, totalPoint: 0, users: new Set(), dates: new Map() });
      }
      const item = map.get(memo);
      item.totalPoint += point;
      item.users.add(userId);
      
      if (!item.dates.has(date)) {
        item.dates.set(date, { point: 0, count: 0 });
      }
      item.dates.get(date).point += point;
      item.dates.get(date).count += 1;
    });
    
    return Array.from(map.values())
      .map(item => ({
        ...item,
        userCount: item.users.size,
        dateList: Array.from(item.dates.entries())
          .map(([date, data]) => ({ date, ...data }))
          .sort((a, b) => a.date.localeCompare(b.date))
      }))
      .sort((a, b) => b.totalPoint - a.totalPoint);
  }, [earnData]);

  const companyStats = useMemo(() => {
    const map = new Map();
    useData.forEach(row => {
      const company = row['업체명'] || '(없음)';
      const point = Number(row['포인트']) || 0;
      const userId = row['고객ID'];
      
      if (!map.has(company)) {
        map.set(company, { company, usedPoint: 0, users: new Set() });
      }
      const c = map.get(company);
      c.usedPoint += point;
      c.users.add(userId);
    });
    return Array.from(map.values())
      .map(c => ({ ...c, userCount: c.users.size }))
      .sort((a, b) => a.usedPoint - b.usedPoint);
  }, [useData]);

  const productStats = useMemo(() => {
    const map = new Map();
    useData.forEach(row => {
      const memo = row['사용자메모'] || '';
      const match = memo.match(/상품명\(([^)]+)\)/);
      const product = match ? match[1] : '(알수없음)';
      const point = Number(row['포인트']) || 0;
      const userId = row['고객ID'];
      
      if (!map.has(product)) {
        map.set(product, { product, usedPoint: 0, users: new Set() });
      }
      const p = map.get(product);
      p.usedPoint += point;
      p.users.add(userId);
    });
    return Array.from(map.values())
      .map(p => ({ ...p, userCount: p.users.size }))
      .sort((a, b) => a.usedPoint - b.usedPoint);
  }, [useData]);

  // 사용자별 상세 - 수정된 로직
  const userStats = useMemo(() => {
    const map = new Map();
    
    // 먼저 당월 데이터를 시간순 정렬
    const sortedMonthlyData = [...monthlyData].sort((a, b) => 
      (a['처리일'] || '').localeCompare(b['처리일'] || '')
    );
    
    // 당월 데이터 수집
    sortedMonthlyData.forEach(row => {
      const id = row['고객ID'];
      const name = row['고객명'];
      const point = Number(row['포인트']) || 0;
      const total = Number(row['토탈포인트']) || 0;
      const type = row['타입'];
      const date = row['처리일'] || '';
      
      if (!map.has(id)) {
        map.set(id, { 
          id, name, 
          startPoint: 0,
          earnedPoint: 0,
          usedPoint: 0,
          currentPoint: 0,
          calculatedPoint: 0,
          mismatch: false,
          transactions: [],
          lastDate: ''
          lastOrderNo: '' //추가 
        });
      }
      const user = map.get(id);
      user.name = name;
      
      if (type === '사용') {
        user.usedPoint += point;
      } else {
        user.earnedPoint += point;
      }
      
    // 같은 시간+주문번호일 경우 가장 작은 토탈포인트가 최종
const orderNo = row['주문번호'] || '';
if (date > user.lastDate) {
  user.currentPoint = total;
  user.lastDate = date;
  user.lastOrderNo = orderNo;
} else if (date === user.lastDate) {
  // 동일 시간이면 더 작은 토탈포인트 선택
  if (total < user.currentPoint) {
    user.currentPoint = total;
  }
}
      
      user.transactions.push({ 
        date, type, point, total, 
        memo: row['사용자메모'] || row['관리자메모'] || '',
        status: row['주문상태'] || ''
      });
    });
    
    // 시작 포인트 계산
    if (selectedMonth) {
      // 전월 데이터에서 각 사용자의 마지막 잔액
      const prevData = validData
        .filter(row => {
          const date = row['처리일'] || '';
          const rowMonth = date.substring(0, 7).replace('/', '-');
          return rowMonth < selectedMonth;
        })
        .sort((a, b) => (a['처리일'] || '').localeCompare(b['처리일'] || ''));
      
      const prevUserTotal = new Map();
      prevData.forEach(row => {
        prevUserTotal.set(row['고객ID'], Number(row['토탈포인트']) || 0);
      });
      
      map.forEach((user, id) => {
        if (prevUserTotal.has(id)) {
          user.startPoint = prevUserTotal.get(id);
        } else if (user.transactions.length > 0) {
          // 전월 데이터 없으면 첫 거래에서 역산
          const sortedTx = [...user.transactions].sort((a, b) => a.date.localeCompare(b.date));
          const firstTx = sortedTx[0];
          user.startPoint = firstTx.total - firstTx.point;
        }
        
        user.calculatedPoint = user.startPoint + user.earnedPoint + user.usedPoint;
        user.mismatch = Math.abs(user.calculatedPoint - user.currentPoint) > 1;
      });
    }
    
    return Array.from(map.values())
      .filter(u => u.usedPoint !== 0 || u.earnedPoint !== 0)
      .sort((a, b) => a.usedPoint - b.usedPoint);
  }, [monthlyData, validData, selectedMonth]);

  const mismatchCount = useMemo(() => userStats.filter(u => u.mismatch).length, [userStats]);

  const filteredEarnByType = useMemo(() => {
    if (!searchTerm) return earnByType;
    const lower = searchTerm.toLowerCase();
    return earnByType.filter(e => e.memo.toLowerCase().includes(lower));
  }, [earnByType, searchTerm]);

  const filteredCompanyStats = useMemo(() => {
    if (!searchTerm) return companyStats;
    const lower = searchTerm.toLowerCase();
    return companyStats.filter(c => c.company.toLowerCase().includes(lower));
  }, [companyStats, searchTerm]);

  const filteredProductStats = useMemo(() => {
    if (!searchTerm) return productStats;
    const lower = searchTerm.toLowerCase();
    return productStats.filter(p => p.product.toLowerCase().includes(lower));
  }, [productStats, searchTerm]);

  const filteredUserStats = useMemo(() => {
    if (!searchTerm) return userStats;
    const lower = searchTerm.toLowerCase();
    return userStats.filter(u => u.name?.toLowerCase().includes(lower) || u.id?.toLowerCase().includes(lower));
  }, [userStats, searchTerm]);

  const handleDownload = () => {
    let downloadData = [];
    let sheetName = '';
    const monthLabel = selectedMonth || '전체';
    
    if (mainTab === 'earn') {
      downloadData = filteredEarnByType.map(e => ({
        '적립유형(관리자메모)': e.memo,
        '총적립포인트': e.totalPoint,
        '적립인원': e.userCount
      }));
      sheetName = '적립내역';
    } else if (useSubTab === 'company') {
      downloadData = filteredCompanyStats.map(c => ({
        '업체명': c.company,
        '사용포인트': c.usedPoint,
        '사용인원': c.userCount
      }));
      sheetName = '업체별';
    } else if (useSubTab === 'product') {
      downloadData = filteredProductStats.map(p => ({
        '상품명': p.product,
        '사용포인트': p.usedPoint,
        '사용인원': p.userCount
      }));
      sheetName = '상품별';
    } else {
      downloadData = filteredUserStats.map(u => ({
        '고객ID': u.id,
        '고객명': u.name,
        '시작포인트': u.startPoint,
        '적립포인트': u.earnedPoint,
        '사용포인트': u.usedPoint,
        '계산잔여': u.calculatedPoint,
        '실제잔여': u.currentPoint,
        '불일치': u.mismatch ? 'O' : ''
      }));
      sheetName = '사용자별';
    }
    
    const ws = XLSX.utils.json_to_sheet(downloadData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${monthLabel}_${sheetName}_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  const getMonthLabel = () => {
    if (!selectedMonth) return '';
    const [year, month] = selectedMonth.split('-');
    return `${year}년 ${parseInt(month)}월`;
  };

  const mainTabStyle = (tab) => ({
    padding: '14px 28px',
    border: 'none',
    backgroundColor: mainTab === tab ? (tab === 'earn' ? '#16a34a' : '#dc2626') : '#e5e7eb',
    color: mainTab === tab ? 'white' : '#666',
    cursor: 'pointer',
    fontWeight: 'bold',
    fontSize: '15px',
    borderRadius: tab === 'earn' ? '8px 0 0 8px' : '0 8px 8px 0'
  });

  const subTabStyle = (tab) => ({
    padding: '10px 20px',
    border: 'none',
    borderBottom: useSubTab === tab ? '3px solid #dc2626' : '3px solid transparent',
    backgroundColor: 'transparent',
    cursor: 'pointer',
    fontWeight: useSubTab === tab ? 'bold' : 'normal',
    color: useSubTab === tab ? '#dc2626' : '#666'
  });

  const thStyle = { padding: '12px', textAlign: 'left', borderBottom: '2px solid #e5e7eb', backgroundColor: '#f9fafb', whiteSpace: 'nowrap' };
  const tdStyle = { padding: '12px', borderBottom: '1px solid #e5e7eb' };
  const cardStyle = {
    backgroundColor: 'white',
    borderRadius: '12px',
    padding: '20px',
    textAlign: 'center',
    boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
    flex: 1,
    minWidth: '140px'
  };

  return (
    <div style={{ padding: '16px', backgroundColor: '#f5f5f5', minHeight: '100vh', fontFamily: 'sans-serif' }}>
      <h1 style={{ fontSize: '24px', fontWeight: 'bold', marginBottom: '20px' }}>📊 {getMonthLabel() || ''} 포인트</h1>
      
      <div style={{ backgroundColor: 'white', borderRadius: '8px', padding: '16px', marginBottom: '16px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
        <div style={{ display: 'flex', gap: '12px', alignItems: 'center', flexWrap: 'wrap' }}>
          <label style={{ backgroundColor: '#2563eb', color: 'white', padding: '10px 20px', borderRadius: '6px', cursor: 'pointer' }}>
            📁 엑셀 업로드
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} style={{ display: 'none' }} />
          </label>
          
          {availableMonths.length > 0 && (
            <select 
              value={selectedMonth} 
              onChange={(e) => setSelectedMonth(e.target.value)}
              style={{ padding: '10px 16px', borderRadius: '6px', border: '1px solid #d1d5db', fontSize: '14px' }}
            >
              <option value="">전체 기간</option>
              {availableMonths.map(m => (
                <option key={m} value={m}>{m.replace('-', '년 ')}월</option>
              ))}
            </select>
          )}
          
          {fileName && <span style={{ color: '#666', fontSize: '14px' }}>{fileName}</span>}
        </div>
        
        {data.length > 0 && (
          <div style={{ marginTop: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer' }}>
              <input 
                type="checkbox" 
                checked={showCanceled} 
                onChange={(e) => setShowCanceled(e.target.checked)}
                style={{ width: '18px', height: '18px' }}
              />
              <span style={{ fontSize: '14px' }}>취소완료 포함</span>
            </label>
            <span style={{ fontSize: '13px', color: '#666' }}>
              (취소완료 {canceledCount}건 {showCanceled ? '포함됨' : '제외됨'})
            </span>
          </div>
        )}
      </div>

      {data.length > 0 && (
        <>
          <div style={{ display: 'flex', gap: '12px', marginBottom: '20px', flexWrap: 'wrap' }}>
            <div style={{ ...cardStyle, borderTop: '4px solid #8b5cf6' }}>
              <div style={{ color: '#666', fontSize: '13px', marginBottom: '6px' }}>📦 전월 이월</div>
              <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#8b5cf6' }}>
                {monthlyTotals.carryover.toLocaleString()}
              </div>
            </div>
            <div style={{ ...cardStyle, borderTop: '4px solid #16a34a' }}>
              <div style={{ color: '#666', fontSize: '13px', marginBottom: '6px' }}>➕ 적립 ({earnData.length}건)</div>
              <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#16a34a' }}>
                +{monthlyTotals.earned.toLocaleString()}
              </div>
            </div>
            <div style={{ ...cardStyle, borderTop: '4px solid #dc2626' }}>
              <div style={{ color: '#666', fontSize: '13px', marginBottom: '6px' }}>➖ 사용 ({useData.length}건)</div>
              <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#dc2626' }}>
                {monthlyTotals.used.toLocaleString()}
              </div>
            </div>
            <div style={{ ...cardStyle, borderTop: '4px solid #2563eb', backgroundColor: '#eff6ff' }}>
              <div style={{ color: '#666', fontSize: '13px', marginBottom: '6px' }}>💰 잔여 포인트</div>
              <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#2563eb' }}>
                {monthlyTotals.balance.toLocaleString()}
              </div>
            </div>
          </div>

          <div style={{ marginBottom: '16px', display: 'flex' }}>
            <button style={mainTabStyle('earn')} onClick={() => { setMainTab('earn'); setSearchTerm(''); }}>
              ➕ 적립내역
            </button>
            <button style={mainTabStyle('use')} onClick={() => { setMainTab('use'); setSearchTerm(''); }}>
              ➖ 사용내역
            </button>
          </div>

          {mainTab === 'earn' && (
            <div style={{ backgroundColor: 'white', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', overflow: 'hidden' }}>
              <div style={{ padding: '12px', display: 'flex', gap: '12px', borderBottom: '1px solid #e5e7eb' }}>
                <input
                  type="text"
                  placeholder="적립유형 검색..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  style={{ flex: 1, padding: '10px', border: '1px solid #d1d5db', borderRadius: '6px' }}
                />
                <button onClick={handleDownload} style={{ backgroundColor: '#16a34a', color: 'white', padding: '10px 20px', borderRadius: '6px', border: 'none', cursor: 'pointer' }}>
                  📥 다운로드
                </button>
              </div>
              
              <div style={{ padding: '12px', backgroundColor: '#f0fdf4', borderBottom: '1px solid #e5e7eb' }}>
                <strong>적립유형 {filteredEarnByType.length}개 / 총 {earnData.length}건 / {new Set(earnData.map(e => e['고객ID'])).size}명</strong>
              </div>
              
              <div style={{ overflowX: 'auto', maxHeight: '500px', overflowY: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
                  <thead style={{ position: 'sticky', top: 0 }}>
                    <tr>
                      <th style={{ ...thStyle, width: '40px' }}></th>
                      <th style={thStyle}>적립유형 (관리자메모)</th>
                      <th style={{ ...thStyle, textAlign: 'right' }}>총 적립포인트</th>
                      <th style={{ ...thStyle, textAlign: 'right' }}>적립인원</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredEarnByType.map((e, i) => (
                      <React.Fragment key={i}>
                        <tr 
                          style={{ cursor: 'pointer', backgroundColor: expandedEarn === i ? '#f0fdf4' : 'transparent' }}
                          onClick={() => setExpandedEarn(expandedEarn === i ? null : i)}
                        >
                          <td style={tdStyle}>{expandedEarn === i ? '▼' : '▶'}</td>
                          <td style={{ ...tdStyle, maxWidth: '400px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={e.memo}>
                            {e.memo}
                          </td>
                          <td style={{ ...tdStyle, textAlign: 'right', color: '#16a34a', fontWeight: 'bold' }}>
                            +{e.totalPoint.toLocaleString()}
                          </td>
                          <td style={{ ...tdStyle, textAlign: 'right' }}>{e.userCount}명</td>
                        </tr>
                        {expandedEarn === i && (
                          <tr>
                            <td colSpan={4} style={{ padding: '0', backgroundColor: '#f9fafb' }}>
                              <div style={{ padding: '12px 20px' }}>
                                <div style={{ fontWeight: 'bold', marginBottom: '8px', color: '#666' }}>📅 일자별 상세</div>
                                <table style={{ width: '100%', fontSize: '13px' }}>
                                  <thead>
                                    <tr style={{ backgroundColor: '#e5e7eb' }}>
                                      <th style={{ padding: '8px', textAlign: 'left' }}>일자</th>
                                      <th style={{ padding: '8px', textAlign: 'right' }}>적립포인트</th>
                                      <th style={{ padding: '8px', textAlign: 'right' }}>건수</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {e.dateList.map((d, j) => (
                                      <tr key={j}>
                                        <td style={{ padding: '8px' }}>{d.date}</td>
                                        <td style={{ padding: '8px', textAlign: 'right', color: '#16a34a' }}>+{d.point.toLocaleString()}</td>
                                        <td style={{ padding: '8px', textAlign: 'right' }}>{d.count}건</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {mainTab === 'use' && (
            <div style={{ backgroundColor: 'white', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', overflow: 'hidden' }}>
              <div style={{ display: 'flex', borderBottom: '1px solid #e5e7eb' }}>
                <button style={subTabStyle('company')} onClick={() => { setUseSubTab('company'); setSearchTerm(''); }}>🏢 업체별</button>
                <button style={subTabStyle('product')} onClick={() => { setUseSubTab('product'); setSearchTerm(''); }}>📦 상품별</button>
                <button style={subTabStyle('user')} onClick={() => { setUseSubTab('user'); setSearchTerm(''); }}>
                  👤 사용자별 {mismatchCount > 0 && <span style={{ color: '#dc2626', marginLeft: '4px' }}>⚠️{mismatchCount}</span>}
                </button>
              </div>

              <div style={{ padding: '12px', display: 'flex', gap: '12px', borderBottom: '1px solid #e5e7eb' }}>
                <input
                  type="text"
                  placeholder={useSubTab === 'company' ? '업체명 검색...' : useSubTab === 'product' ? '상품명 검색...' : '이름 검색...'}
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  style={{ flex: 1, padding: '10px', border: '1px solid #d1d5db', borderRadius: '6px' }}
                />
                <button onClick={handleDownload} style={{ backgroundColor: '#dc2626', color: 'white', padding: '10px 20px', borderRadius: '6px', border: 'none', cursor: 'pointer' }}>
                  📥 다운로드
                </button>
              </div>

              {useSubTab === 'company' && (
                <>
                  <div style={{ padding: '12px', backgroundColor: '#fef2f2', borderBottom: '1px solid #e5e7eb' }}>
                    <strong>총 {filteredCompanyStats.length}개 업체</strong>
                  </div>
                  <div style={{ overflowX: 'auto', maxHeight: '500px', overflowY: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
                      <thead style={{ position: 'sticky', top: 0 }}>
                        <tr>
                          <th style={thStyle}>업체명</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>사용 포인트</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>사용 인원</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredCompanyStats.map((c, i) => (
                          <tr key={i}>
                            <td style={tdStyle}>{c.company}</td>
                            <td style={{ ...tdStyle, textAlign: 'right', color: '#dc2626', fontWeight: 'bold' }}>{c.usedPoint.toLocaleString()}</td>
                            <td style={{ ...tdStyle, textAlign: 'right' }}>{c.userCount}명</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}

              {useSubTab === 'product' && (
                <>
                  <div style={{ padding: '12px', backgroundColor: '#fef2f2', borderBottom: '1px solid #e5e7eb' }}>
                    <strong>총 {filteredProductStats.length}개 상품</strong>
                  </div>
                  <div style={{ overflowX: 'auto', maxHeight: '500px', overflowY: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
                      <thead style={{ position: 'sticky', top: 0 }}>
                        <tr>
                          <th style={thStyle}>상품명</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>사용 포인트</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>사용 인원</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredProductStats.map((p, i) => (
                          <tr key={i}>
                            <td style={{ ...tdStyle, maxWidth: '400px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={p.product}>{p.product}</td>
                            <td style={{ ...tdStyle, textAlign: 'right', color: '#dc2626', fontWeight: 'bold' }}>{p.usedPoint.toLocaleString()}</td>
                            <td style={{ ...tdStyle, textAlign: 'right' }}>{p.userCount}명</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}

              {useSubTab === 'user' && (
                <>
                  <div style={{ padding: '12px', backgroundColor: '#fef2f2', borderBottom: '1px solid #e5e7eb', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <strong>총 {filteredUserStats.length}명</strong>
                    {mismatchCount > 0 && (
                      <span style={{ color: '#dc2626', fontSize: '13px' }}>⚠️ 계산 불일치 {mismatchCount}명</span>
                    )}
                  </div>
                  <div style={{ overflowX: 'auto', maxHeight: '500px', overflowY: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
                      <thead style={{ position: 'sticky', top: 0 }}>
                        <tr>
                          <th style={{ ...thStyle, width: '40px' }}></th>
                          <th style={thStyle}>이름</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>시작</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>적립</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>사용</th>
                          <th style={{ ...thStyle, textAlign: 'right' }}>= 잔여</th>
                          <th style={{ ...thStyle, textAlign: 'center' }}>검증</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredUserStats.map((u, i) => (
                          <React.Fragment key={i}>
                            <tr 
                              style={{ 
                                cursor: 'pointer', 
                                backgroundColor: u.mismatch ? '#fef2f2' : (expandedUser === i ? '#f0f9ff' : 'transparent')
                              }}
                              onClick={() => setExpandedUser(expandedUser === i ? null : i)}
                            >
                              <td style={tdStyle}>{expandedUser === i ? '▼' : '▶'}</td>
                              <td style={{ ...tdStyle, fontWeight: '500' }}>{u.name}</td>
                              <td style={{ ...tdStyle, textAlign: 'right', color: '#8b5cf6' }}>{u.startPoint.toLocaleString()}</td>
                              <td style={{ ...tdStyle, textAlign: 'right', color: '#16a34a' }}>+{u.earnedPoint.toLocaleString()}</td>
                              <td style={{ ...tdStyle, textAlign: 'right', color: '#dc2626', fontWeight: 'bold' }}>{u.usedPoint.toLocaleString()}</td>
                              <td style={{ ...tdStyle, textAlign: 'right', fontWeight: 'bold', color: '#2563eb' }}>{u.currentPoint.toLocaleString()}</td>
                              <td style={{ ...tdStyle, textAlign: 'center' }}>
                                {u.mismatch ? (
                                  <span style={{ color: '#dc2626' }}>⚠️ {u.calculatedPoint.toLocaleString()}</span>
                                ) : (
                                  <span style={{ color: '#16a34a' }}>✓</span>
                                )}
                              </td>
                            </tr>
                            {expandedUser === i && (
                              <tr>
                                <td colSpan={7} style={{ padding: '0', backgroundColor: '#f9fafb' }}>
                                  <div style={{ padding: '12px 20px' }}>
                                    <div style={{ fontWeight: 'bold', marginBottom: '8px', color: '#666' }}>📋 거래 내역</div>
                                    <div style={{ fontSize: '12px', color: '#666', marginBottom: '8px', padding: '8px', backgroundColor: '#e5e7eb', borderRadius: '4px' }}>
                                      시작 <strong>{u.startPoint.toLocaleString()}</strong> + 
                                      적립 <strong style={{ color: '#16a34a' }}>+{u.earnedPoint.toLocaleString()}</strong> + 
                                      사용 <strong style={{ color: '#dc2626' }}>{u.usedPoint.toLocaleString()}</strong> = 
                                      계산 <strong>{u.calculatedPoint.toLocaleString()}</strong> / 
                                      실제 <strong style={{ color: '#2563eb' }}>{u.currentPoint.toLocaleString()}</strong>
                                      {u.mismatch && <span style={{ color: '#dc2626', marginLeft: '8px' }}>⚠️ 차이: {(u.currentPoint - u.calculatedPoint).toLocaleString()}</span>}
                                    </div>
                                    <table style={{ width: '100%', fontSize: '13px' }}>
                                      <thead>
                                        <tr style={{ backgroundColor: '#e5e7eb' }}>
                                          <th style={{ padding: '8px', textAlign: 'left' }}>일시</th>
                                          <th style={{ padding: '8px', textAlign: 'left' }}>타입</th>
                                          <th style={{ padding: '8px', textAlign: 'right' }}>포인트</th>
                                          <th style={{ padding: '8px', textAlign: 'right' }}>잔액</th>
                                          <th style={{ padding: '8px', textAlign: 'left' }}>상태</th>
                                          <th style={{ padding: '8px', textAlign: 'left' }}>메모</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {[...u.transactions].sort((a, b) => a.date.localeCompare(b.date)).map((tx, j) => (
                                          <tr key={j}>
                                            <td style={{ padding: '8px' }}>{tx.date}</td>
                                            <td style={{ padding: '8px' }}>
                                              <span style={{ 
                                                padding: '2px 6px', borderRadius: '4px', fontSize: '11px',
                                                backgroundColor: tx.type === '사용' ? '#fef2f2' : '#f0fdf4',
                                                color: tx.type === '사용' ? '#dc2626' : '#16a34a'
                                              }}>
                                                {tx.type}
                                              </span>
                                            </td>
                                            <td style={{ padding: '8px', textAlign: 'right', color: tx.point < 0 ? '#dc2626' : '#16a34a', fontWeight: '500' }}>
                                              {tx.point > 0 ? '+' : ''}{tx.point.toLocaleString()}
                                            </td>
                                            <td style={{ padding: '8px', textAlign: 'right' }}>{tx.total.toLocaleString()}</td>
                                            <td style={{ padding: '8px', fontSize: '12px' }}>{tx.status || '-'}</td>
                                            <td style={{ padding: '8px', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: '12px', color: '#666' }} title={tx.memo}>
                                              {tx.memo || '-'}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}
        </>
      )}

      {data.length === 0 && (
        <div style={{ backgroundColor: 'white', borderRadius: '8px', padding: '60px', textAlign: 'center', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
          <div style={{ fontSize: '48px', marginBottom: '16px' }}>📄</div>
          <p style={{ color: '#666', marginBottom: '8px' }}>엑셀 파일을 업로드하세요</p>
          <p style={{ color: '#999', fontSize: '14px' }}>.xlsx, .xls, .csv 지원</p>
        </div>
      )}
    </div>
  );
}
