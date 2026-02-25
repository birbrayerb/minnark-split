import { useState, useCallback } from 'react';
import { Upload, FileSpreadsheet, CheckCircle2, AlertCircle, Download, ClipboardCopy, ChevronDown, ChevronUp, BarChart3, TrendingUp, X } from 'lucide-react';
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import { readFile, detectFileType, processFiles, exportCSV, exportExcel } from './pipeline';

const fmt = (n) => n.toLocaleString('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2 });
const fmtShort = (n) => {
  if (Math.abs(n) >= 1e6) return `$${(n/1e6).toFixed(1)}M`;
  if (Math.abs(n) >= 1e3) return `$${(n/1e3).toFixed(1)}K`;
  return fmt(n);
};

const COLORS = {
  blackfin: '#3b82f6',
  blackfinLight: '#60a5fa',
  mizar: '#14b8a6',
  mizarLight: '#2dd4bf',
  di: '#8b5cf6',
  domestic: '#f59e0b',
};

function App() {
  const [files, setFiles] = useState([]);
  const [results, setResults] = useState(null);
  const [error, setError] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [sortCol, setSortCol] = useState('program');
  const [sortDir, setSortDir] = useState('asc');
  const [filterTeam, setFilterTeam] = useState('all');
  const [filterCat, setFilterCat] = useState('all');
  const [copied, setCopied] = useState(false);

  const handleFiles = useCallback(async (newFiles) => {
    setError(null);
    const fileList = Array.from(newFiles).filter(f =>
      f.name.endsWith('.xlsx') || f.name.endsWith('.xls')
    );
    if (fileList.length === 0) {
      setError('Please upload Excel files (.xlsx)');
      return;
    }

    const processed = [];
    for (const f of fileList) {
      try {
        const wb = await readFile(f);
        const type = detectFileType(wb);
        processed.push({ file: f, workbook: wb, type, name: f.name, status: type ? 'ready' : 'unknown' });
      } catch (err) {
        processed.push({ file: f, name: f.name, type: null, status: 'error', error: err.message });
      }
    }
    setFiles(prev => [...prev, ...processed]);
  }, []);

  const removeFile = (idx) => {
    setFiles(prev => prev.filter((_, i) => i !== idx));
    setResults(null);
  };

  const runProcessing = useCallback(async () => {
    setProcessing(true);
    setError(null);
    try {
      const jpmFiles = files.filter(f => f.type === 'jpm').map(f => ({ workbook: f.workbook, filename: f.name }));
      const domesticFile = files.find(f => f.type === 'domestic');
      if (jpmFiles.length === 0) throw new Error('No JPM monthly files detected. Upload files with TOTAL DI / SUMMARY sheets.');
      const r = processFiles(jpmFiles, domesticFile ? { workbook: domesticFile.workbook, filename: domesticFile.name } : null);
      if (r.length === 0) throw new Error('No data found in uploaded files.');
      setResults(r);
    } catch (err) {
      setError(err.message);
    }
    setProcessing(false);
  }, [files]);

  const handleDrop = (e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); };
  const handleDragOver = (e) => { e.preventDefault(); setDragOver(true); };
  const handleDragLeave = () => setDragOver(false);

  const handleExportCSV = () => {
    const csv = exportCSV(results);
    downloadBlob(new Blob([csv], { type: 'text/csv' }), 'minnark-split.csv');
  };

  const handleExportExcel = () => {
    const buf = exportExcel(results);
    downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'minnark-split.xlsx');
  };

  const handleCopy = () => {
    const csv = exportCSV(results);
    navigator.clipboard.writeText(csv);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const downloadBlob = (blob, name) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name; a.click();
    URL.revokeObjectURL(url);
  };

  // Build detail table data
  const detailRows = results ? results.flatMap(r => {
    const rows = [];
    const add = (cat, team, progs) => {
      for (const [prog, amt] of Object.entries(progs)) {
        rows.push({ month: r.month, category: cat, team, program: prog, amount: amt });
      }
    };
    add('DI', 'Blackfin', r.di_blackfin_programs);
    add('DI', 'Mizar', r.di_mizar_programs);
    add('DO', 'Blackfin', r.do_blackfin_programs);
    add('DO', 'Mizar', r.do_mizar_programs);
    return rows;
  }) : [];

  const filtered = detailRows.filter(r =>
    (filterTeam === 'all' || r.team === filterTeam) &&
    (filterCat === 'all' || r.category === filterCat)
  );

  const sorted = [...filtered].sort((a, b) => {
    const mul = sortDir === 'asc' ? 1 : -1;
    if (sortCol === 'amount') return (a.amount - b.amount) * mul;
    return String(a[sortCol]).localeCompare(String(b[sortCol])) * mul;
  });

  const toggleSort = (col) => {
    if (sortCol === col) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortCol(col); setSortDir('asc'); }
  };

  const SortIcon = ({ col }) => sortCol === col
    ? (sortDir === 'asc' ? <ChevronUp className="inline w-3 h-3" /> : <ChevronDown className="inline w-3 h-3" />)
    : null;

  // Aggregate totals for charts
  const totals = results ? results.reduce((acc, r) => ({
    di_blackfin: acc.di_blackfin + r.actual.di_blackfin,
    di_mizar: acc.di_mizar + r.actual.di_mizar,
    do_blackfin: acc.do_blackfin + r.actual.do_blackfin,
    do_mizar: acc.do_mizar + r.actual.do_mizar,
  }), { di_blackfin: 0, di_mizar: 0, do_blackfin: 0, do_mizar: 0 }) : null;

  const grandTotal = totals ? totals.di_blackfin + totals.di_mizar + totals.do_blackfin + totals.do_mizar : 0;
  const blackfinTotal = totals ? totals.di_blackfin + totals.do_blackfin : 0;
  const mizarTotal = totals ? totals.di_mizar + totals.do_mizar : 0;

  return (
    <div className="min-h-screen px-4 py-8 max-w-7xl mx-auto">
      {/* Header */}
      <header className="mb-8 fade-in">
        <div className="flex items-center gap-3 mb-2">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-500 to-teal-500 flex items-center justify-center">
            <BarChart3 className="w-5 h-5 text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-white tracking-tight">MinnARK Revenue Split</h1>
            <p className="text-sm text-slate-400">Automated partner revenue allocation</p>
          </div>
        </div>
      </header>

      {/* Upload Section */}
      <section className="mb-8 fade-in">
        <div
          className={`upload-zone p-8 text-center ${dragOver ? 'drag-over' : ''}`}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onClick={() => document.getElementById('file-input').click()}
        >
          <Upload className="w-10 h-10 mx-auto mb-3 text-slate-400" />
          <p className="text-slate-300 font-medium">Drop Excel files here or click to upload</p>
          <p className="text-xs text-slate-500 mt-1">JPM monthly files • Domestic Payments file</p>
          <input id="file-input" type="file" multiple accept=".xlsx,.xls" className="hidden"
            onChange={(e) => { handleFiles(e.target.files); e.target.value = ''; }} />
        </div>

        {files.length > 0 && (
          <div className="mt-4 space-y-2">
            {files.map((f, i) => (
              <div key={i} className="glass-card-sm flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <FileSpreadsheet className="w-5 h-5 text-slate-400" />
                  <div>
                    <p className="text-sm font-medium text-slate-200">{f.name}</p>
                    <p className="text-xs text-slate-500">
                      {f.type === 'jpm' ? '📊 JPM Monthly' : f.type === 'domestic' ? '📋 Domestic Payments' : '❓ Unknown format'}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  {f.status === 'ready' && <CheckCircle2 className="w-4 h-4 text-emerald-400" />}
                  {f.status === 'error' && <AlertCircle className="w-4 h-4 text-red-400" />}
                  {f.status === 'unknown' && <AlertCircle className="w-4 h-4 text-amber-400" />}
                  <button onClick={() => removeFile(i)} className="p-1 hover:bg-slate-700 rounded">
                    <X className="w-4 h-4 text-slate-500" />
                  </button>
                </div>
              </div>
            ))}
            <button
              onClick={runProcessing}
              disabled={processing || files.filter(f => f.type === 'jpm').length === 0}
              className="mt-3 px-6 py-2.5 bg-gradient-to-r from-blue-600 to-teal-600 text-white font-medium rounded-xl
                hover:from-blue-500 hover:to-teal-500 disabled:opacity-40 disabled:cursor-not-allowed transition-all"
            >
              {processing ? (
                <span className="flex items-center gap-2"><span className="pulse-soft">Processing...</span></span>
              ) : (
                <span className="flex items-center gap-2"><TrendingUp className="w-4 h-4" /> Process Revenue Split</span>
              )}
            </button>
          </div>
        )}
      </section>

      {error && (
        <div className="mb-6 p-4 rounded-xl bg-red-500/10 border border-red-500/30 text-red-300 text-sm fade-in">
          <AlertCircle className="w-4 h-4 inline mr-2" />{error}
        </div>
      )}

      {/* Results */}
      {results && (
        <div className="space-y-6 fade-in">
          {/* Grand Total */}
          <div className="glass-card text-center">
            <p className="text-sm text-slate-400 uppercase tracking-wider mb-1">Total Revenue</p>
            <p className="text-4xl font-bold text-white">{fmt(grandTotal)}</p>
            <div className="flex justify-center gap-8 mt-4">
              <div>
                <div className="w-3 h-3 rounded-full inline-block mr-2" style={{ background: COLORS.blackfin }} />
                <span className="text-sm text-slate-300">Blackfin: {fmt(blackfinTotal)}</span>
                <span className="text-xs text-slate-500 ml-1">({(blackfinTotal/grandTotal*100).toFixed(1)}%)</span>
              </div>
              <div>
                <div className="w-3 h-3 rounded-full inline-block mr-2" style={{ background: COLORS.mizar }} />
                <span className="text-sm text-slate-300">Mizar: {fmt(mizarTotal)}</span>
                <span className="text-xs text-slate-500 ml-1">({(mizarTotal/grandTotal*100).toFixed(1)}%)</span>
              </div>
            </div>
            {/* Split bar */}
            <div className="mt-4 h-3 rounded-full overflow-hidden bg-slate-700 max-w-md mx-auto">
              <div className="h-full rounded-full" style={{
                width: `${blackfinTotal/grandTotal*100}%`,
                background: `linear-gradient(to right, ${COLORS.blackfin}, ${COLORS.blackfinLight})`
              }} />
            </div>
          </div>

          {/* Monthly Cards + Charts */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Monthly breakdown cards */}
            <div className="space-y-4">
              <h2 className="text-lg font-semibold text-white">Monthly Breakdown</h2>
              {results.map((r, i) => (
                <div key={i} className="glass-card">
                  <h3 className="text-base font-semibold text-white mb-3">{r.month}</h3>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-xs text-slate-400 uppercase mb-2">Direct Import</p>
                      <div className="space-y-2">
                        <div className="flex justify-between text-sm">
                          <span style={{ color: COLORS.blackfin }}>Blackfin</span>
                          <span className="text-slate-200">{fmt(r.actual.di_blackfin)}</span>
                        </div>
                        <div className="flex justify-between text-sm">
                          <span style={{ color: COLORS.mizar }}>Mizar</span>
                          <span className="text-slate-200">{fmt(r.actual.di_mizar)}</span>
                        </div>
                        <div className="flex justify-between text-sm font-medium border-t border-slate-600 pt-1">
                          <span className="text-slate-400">Total</span>
                          <span className="text-white">{fmt(r.actual.di_total)}</span>
                        </div>
                      </div>
                    </div>
                    <div>
                      <p className="text-xs text-slate-400 uppercase mb-2">Domestic</p>
                      <div className="space-y-2">
                        <div className="flex justify-between text-sm">
                          <span style={{ color: COLORS.blackfin }}>Blackfin</span>
                          <span className="text-slate-200">{fmt(r.actual.do_blackfin)}</span>
                        </div>
                        <div className="flex justify-between text-sm">
                          <span style={{ color: COLORS.mizar }}>Mizar</span>
                          <span className="text-slate-200">{fmt(r.actual.do_mizar)}</span>
                        </div>
                        <div className="flex justify-between text-sm font-medium border-t border-slate-600 pt-1">
                          <span className="text-slate-400">Total</span>
                          <span className="text-white">{fmt(r.actual.do_total)}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="mt-3 pt-3 border-t border-slate-600 flex justify-between text-sm font-semibold">
                    <span className="text-slate-300">Grand Total</span>
                    <span className="text-white">{fmt(r.actual.grand_total)}</span>
                  </div>
                </div>
              ))}
            </div>

            {/* Charts */}
            <div className="space-y-4">
              <h2 className="text-lg font-semibold text-white">Visualizations</h2>

              {/* Pie chart - team split */}
              <div className="glass-card">
                <h3 className="text-sm font-medium text-slate-300 mb-4">Partner Split</h3>
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={[
                      { name: 'Blackfin', value: blackfinTotal },
                      { name: 'Mizar', value: mizarTotal },
                    ]} cx="50%" cy="50%" innerRadius={55} outerRadius={85} paddingAngle={3} dataKey="value">
                      <Cell fill={COLORS.blackfin} />
                      <Cell fill={COLORS.mizar} />
                    </Pie>
                    <Tooltip formatter={(v) => fmt(v)} contentStyle={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, color: '#e2e8f0' }} />
                    <Legend formatter={(v) => <span style={{ color: '#cbd5e1', fontSize: 12 }}>{v}</span>} />
                  </PieChart>
                </ResponsiveContainer>
              </div>

              {/* Bar chart - monthly comparison */}
              {results.length > 0 && (
                <div className="glass-card">
                  <h3 className="text-sm font-medium text-slate-300 mb-4">Monthly Revenue by Partner</h3>
                  <ResponsiveContainer width="100%" height={250}>
                    <BarChart data={results.map(r => ({
                      month: r.month.split(' ')[0],
                      Blackfin: r.actual.di_blackfin + r.actual.do_blackfin,
                      Mizar: r.actual.di_mizar + r.actual.do_mizar,
                    }))}>
                      <XAxis dataKey="month" tick={{ fill: '#94a3b8', fontSize: 12 }} />
                      <YAxis tickFormatter={fmtShort} tick={{ fill: '#94a3b8', fontSize: 11 }} />
                      <Tooltip formatter={(v) => fmt(v)} contentStyle={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, color: '#e2e8f0' }} />
                      <Legend formatter={(v) => <span style={{ color: '#cbd5e1', fontSize: 12 }}>{v}</span>} />
                      <Bar dataKey="Blackfin" fill={COLORS.blackfin} radius={[4, 4, 0, 0]} />
                      <Bar dataKey="Mizar" fill={COLORS.mizar} radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              )}

              {/* DI vs DO chart */}
              <div className="glass-card">
                <h3 className="text-sm font-medium text-slate-300 mb-4">DI vs Domestic</h3>
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={[
                      { name: 'Direct Import', value: totals.di_blackfin + totals.di_mizar },
                      { name: 'Domestic', value: totals.do_blackfin + totals.do_mizar },
                    ]} cx="50%" cy="50%" innerRadius={55} outerRadius={85} paddingAngle={3} dataKey="value">
                      <Cell fill={COLORS.di} />
                      <Cell fill={COLORS.domestic} />
                    </Pie>
                    <Tooltip formatter={(v) => fmt(v)} contentStyle={{ background: '#1e293b', border: '1px solid #334155', borderRadius: 8, color: '#e2e8f0' }} />
                    <Legend formatter={(v) => <span style={{ color: '#cbd5e1', fontSize: 12 }}>{v}</span>} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {/* Detail Table */}
          <div className="glass-card">
            <div className="flex flex-wrap items-center justify-between gap-4 mb-4">
              <h2 className="text-lg font-semibold text-white">Program Detail</h2>
              <div className="flex gap-2">
                <select value={filterTeam} onChange={e => setFilterTeam(e.target.value)}
                  className="bg-slate-700 text-sm text-slate-200 rounded-lg px-3 py-1.5 border border-slate-600">
                  <option value="all">All Teams</option>
                  <option value="Blackfin">Blackfin</option>
                  <option value="Mizar">Mizar</option>
                </select>
                <select value={filterCat} onChange={e => setFilterCat(e.target.value)}
                  className="bg-slate-700 text-sm text-slate-200 rounded-lg px-3 py-1.5 border border-slate-600">
                  <option value="all">All Categories</option>
                  <option value="DI">Direct Import</option>
                  <option value="DO">Domestic</option>
                </select>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="data-table">
                <thead>
                  <tr>
                    <th onClick={() => toggleSort('month')}>Month <SortIcon col="month" /></th>
                    <th onClick={() => toggleSort('category')}>Category <SortIcon col="category" /></th>
                    <th onClick={() => toggleSort('team')}>Team <SortIcon col="team" /></th>
                    <th onClick={() => toggleSort('program')}>Program <SortIcon col="program" /></th>
                    <th onClick={() => toggleSort('amount')} className="text-right">Amount <SortIcon col="amount" /></th>
                  </tr>
                </thead>
                <tbody>
                  {sorted.map((r, i) => (
                    <tr key={i}>
                      <td>{r.month}</td>
                      <td>
                        <span className={`inline-block px-2 py-0.5 rounded text-xs font-medium ${
                          r.category === 'DI' ? 'bg-purple-500/20 text-purple-300' : 'bg-amber-500/20 text-amber-300'
                        }`}>{r.category === 'DI' ? 'Direct Import' : 'Domestic'}</span>
                      </td>
                      <td>
                        <span className="flex items-center gap-1.5">
                          <span className="w-2 h-2 rounded-full" style={{ background: r.team === 'Blackfin' ? COLORS.blackfin : COLORS.mizar }} />
                          {r.team}
                        </span>
                      </td>
                      <td className="text-slate-300">{r.program}</td>
                      <td className="text-right font-mono text-slate-200">{fmt(r.amount)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr>
                    <td colSpan={4} className="font-semibold text-slate-300">Total ({sorted.length} rows)</td>
                    <td className="text-right font-mono font-semibold text-white">
                      {fmt(sorted.reduce((s, r) => s + r.amount, 0))}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>

          {/* Export buttons */}
          <div className="flex flex-wrap gap-3">
            <button onClick={handleExportCSV} className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-sm transition-colors">
              <Download className="w-4 h-4" /> Export CSV
            </button>
            <button onClick={handleExportExcel} className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-sm transition-colors">
              <FileSpreadsheet className="w-4 h-4" /> Export Excel
            </button>
            <button onClick={handleCopy} className="flex items-center gap-2 px-4 py-2 bg-slate-700 hover:bg-slate-600 text-slate-200 rounded-xl text-sm transition-colors">
              <ClipboardCopy className="w-4 h-4" /> {copied ? 'Copied!' : 'Copy to Clipboard'}
            </button>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="mt-12 text-center text-xs text-slate-600">
        MinnARK Revenue Split Tool • All processing happens locally in your browser
      </footer>
    </div>
  );
}

export default App;
